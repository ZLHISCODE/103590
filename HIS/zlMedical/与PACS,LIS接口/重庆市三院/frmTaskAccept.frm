VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmTaskAccept 
   Caption         =   "���ܼ�����"
   ClientHeight    =   5205
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   8775
   Icon            =   "frmTaskAccept.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6615
      Top             =   2025
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3060
      Left            =   360
      TabIndex        =   0
      Top             =   1110
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   5398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��쵥"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��λ����"
         Object.Width           =   3704
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3900
      Top             =   3240
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
            Picture         =   "frmTaskAccept.frx":6852
            Key             =   "package"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":D0B4
            Key             =   "package_ok"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   7995
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":13916
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":13B36
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":13D56
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":144D0
            Key             =   "Send"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   8775
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   1138
         ButtonWidth     =   1455
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  ����  "
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "  ���� "
               ImageKey        =   "Send"
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ָ��"
               Key             =   "ָ��"
               Object.ToolTipText     =   "ָ��"
               Object.Tag             =   "ָ��"
               ImageKey        =   "Search"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1320
      Left            =   3195
      TabIndex        =   3
      Top             =   810
      Width           =   3135
      _cx             =   5530
      _cy             =   2328
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   4845
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTaskAccept.frx":14C4A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10398
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin VB.Image imgY 
      Height          =   4680
      Left            =   3075
      MousePointer    =   9  'Size W E
      Top             =   660
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileParam 
         Caption         =   "��������(&P)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "����(&E)"
      Begin VB.Menu mnuEditAutoAccept 
         Caption         =   "�����Զ����ܷ���(&R)"
      End
      Begin VB.Menu mnuEditAcceptPerson 
         Caption         =   "����ָ����Ա����(&A)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditResetPerson 
         Caption         =   "��������ѽ�����(&C)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&T)"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)"
      End
   End
End
Attribute VB_Name = "frmTaskAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcnnSQLServer As New ADODB.Connection
Private mstrSQL As String
Private mblnStartUp As Boolean
Private mstrKey As String

Private Enum mCol
    ��쵥��
    ����
    �Ա�
    ���֤��
    �����
    �Ǽ�id
    ������λ
    ��Ŀ = 0
    ���
    ��־
    �ο�
End Enum

Private mlngCount As Long
Private mlngTotal As Long

Private Function ReadUnit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim objItem As ListItem
    
    On Error GoTo errHand
    
    lvw.ListItems.Clear
    
    Set objItem = lvw.ListItems.Add(, "_0", "--------", 1, 1)
    objItem.SubItems(1) = "[��������嵥]"
    
    mstrSQL = "Select A.ID,A.����,B.���� From ���ǼǼ�¼ A,��Լ��λ B Where A.�Ƿ�����=1 AND A.���״̬=4 AND B.ID=A.��Լ��λid"
    Set rs = OpenRecord(rs, mstrSQL, gstrSysName)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            Set objItem = lvw.ListItems.Add(, "_" & rs("ID").Value, rs("����").Value, 2, 2)
            objItem.SubItems(1) = rs("����").Value
    
            rs.MoveNext
        Loop
    End If
        
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
End Function

Private Function ReadItems(ByVal lng�Ǽ�id As Long, ByVal lng����id As Long) As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
'    vsfItem.Rows = 2
'    vsfItem.RowData(1) = 0
'    vsfItem.Cell(flexcpText, 1, 0, 1, vsfItem.Cols - 1) = ""

    mstrSQL = ""

'    rs.Open mstrSQL, gcnOracle
'    If rs.BOF = False Then
'
'        Call LoadGrid(vsf, rs)
'
'    End If
    
    ReadItems = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
    
End Function

Private Function ReadPerson(ByVal lng�Ǽ�id As Long) As Boolean
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    
    If lng�Ǽ�id = 0 Then
        '����
        
        mstrSQL = "Select B.����id As ID,C.���� As ��쵥��,A.�Ǽ�id,B.����,B.�Ա�,B.���֤��,B.�����,B.������λ From �����Ա���� A,������Ϣ B,���ǼǼ�¼ C Where A.����id=B.����id AND C.ID=A.�Ǽ�id AND C.���״̬=4 AND a.��챨��=1 AND C.��Լ��λid Is Null"
        
    Else
        
        '����
        mstrSQL = "Select B.����id As ID,C.���� As ��쵥��,A.�Ǽ�id,B.����,B.�Ա�,B.���֤��,B.�����,B.������λ From �����Ա���� A,������Ϣ B,���ǼǼ�¼ C Where a.��챨��=1 and A.����id=B.����id AND C.ID=A.�Ǽ�id AND C.ID=" & lng�Ǽ�id
        
    End If
    rs.Open mstrSQL, gcnOracle
    If rs.BOF = False Then
         
        Call LoadGrid(vsf, rs)
        
    End If
    
    ReadPerson = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
    
End Function

Private Function ConnectSQLServer(ByVal strSvr As String, ByVal strDb As String, ByVal strUser As String, ByVal strPsw As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    If mcnnSQLServer.State = adStateOpen Then mcnnSQLServer.Close
    mcnnSQLServer.Open "Provider=SQLOLEDB.1;Password=" & strPsw & ";Persist Security Info=True;User ID=" & strUser & ";Initial Catalog=" & strDb & ";Data Source=" & strSvr
    If mcnnSQLServer.State <> adStateOpen Then
        
        ShowSimpleMsg "���ӵ�LIS������ʧ�ܣ�"
        
        Exit Function
    End If
    
    ConnectSQLServer = True
    Exit Function
errHand:
    ShowSimpleMsg Err.Description
End Function

Private Function AcceptResult() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngTotal As Long
    Dim lngCount As Long
    
    On Error GoTo errHand
    
    frmWait.OpenWait Me, "���ܼ�������", True
    
    mstrSQL = "Select A.�Ǽ�id,A.����id,B.���� From �����Ա���� A,������Ϣ B Where A.����id=B.����id AND A.���״̬=4"
    rs.Open mstrSQL, gcnOracle
    If rs.BOF = False Then
        
        lngTotal = rs.RecordCount
        For lngCount = 1 To lngTotal
            
            frmWait.WaitInfo = "���ڽ��ܡ�" & NVL(rs("����")) & "���ļ�������..."
            frmWait.WaitProgress = Format(100 * lngCount / lngTotal, "0.00")
            
            Call AcceptOneResult(NVL(rs("�Ǽ�id"), 0), NVL(rs("����id"), 0))
            rs.MoveNext
        Next
        
        frmWait.CloseWait
    End If
    
    frmWait.CloseWait
    AcceptResult = True
    
    Exit Function
    
errHand:
    frmWait.CloseWait
    ShowSimpleMsg Err.Description
End Function

Private Function ClearOneResult(ByVal lng�Ǽ�id As Long, ByVal lng����id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim str��쵥�� As String
    Dim lng����� As Long
    
    '����Ƿ��Ѿ����ܣ�ֻ����δ���ܵģ�ͨ���걾���ж�
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    
    mstrSQL = "Select C.����,B.����� From �����Ա���� A,������Ϣ B,���ǼǼ�¼ C Where A.����id=B.����id AND C.ID=A.�Ǽ�id"
    If rsTmp.State = adStateOpen Then rsTmp.Close
    rsTmp.Open mstrSQL, gcnOracle
    If rsTmp.BOF Then Exit Function
    
    str��쵥�� = UCase(NVL(rsTmp("����")))
    lng����� = NVL(rsTmp("�����"), 0)
            
    '���ԭ�еļ�����
    mstrSQL = "ZL_ZLLIS_������('" & str��쵥�� & "'," & lng����id & ")"
    gcnOracle.Execute mstrSQL, , adCmdStoredProc

    gcnOracle.CommitTrans
    
    ClearOneResult = True
    
    ShowSimpleMsg "�Ѿ�����˴����ѽ��ܵļ������ݣ�"
    Exit Function
    
errHand:
    
    ShowSimpleMsg Err.Description
    gcnOracle.RollbackTrans
    
End Function

Private Function AcceptOneResult(ByVal lng�Ǽ�id As Long, ByVal lng����id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim str��쵥�� As String
    Dim lng����� As Long
    Dim strCode As String
    Dim lng����id As Long
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim lng�����Ŀid As Long
    
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    
    mstrSQL = "Select C.����,B.����� " & _
            "From �����Ա���� A,������Ϣ B,���ǼǼ�¼ C Where A.����id=B.����id AND C.ID=A.�Ǽ�id AND C.ID=" & lng�Ǽ�id & " AND A.����id=" & lng����id
    If rsTmp.State = adStateOpen Then rsTmp.Close
    rsTmp.Open mstrSQL, gcnOracle
    If rsTmp.BOF Then Exit Function
    
    str��쵥�� = UCase(NVL(rsTmp("����")))
    lng����� = NVL(rsTmp("�����"), 0)
    
    '��д�µļ�����
    mstrSQL = "Select ������Ŀ����,��Ŀ����,��λ,��ֵ�Խ��,�ַ��Խ��,���뷶Χ,�쳣���,������ " & _
                "From Lis_Value Where ��쵥��='" & str��쵥�� & "' AND ����id='" & lng����� & "'"
                
    rs.Open mstrSQL, mcnnSQLServer
    If rs.BOF = False Then
        Do While Not rs.EOF
            
'            If NVL(rs("������Ŀ����")) <> "" Then
'                strCode = "''"
'                varAry = Split(NVL(rs("������Ŀ����")), ",")
'                For lngLoop = 0 To UBound(varAry)
'                    strCode = strCode & ",'" & varAry(lngLoop) & "'"
'                Next
'            End If
'            strCode = "'" & NVL(rs("������Ŀ����")) & "'"

            lng����id = 0
            
            '1.������id
'            mstrSQL = "Select LIS��ϱ��� From ������ĿĿ¼_LIS Where LIS����='" & NVL(rs("������Ŀ����")) & "'"
'            If rsTmp.State = adStateOpen Then rsTmp.Close
'            rsTmp.Open mstrSQL, gcnOracle
'            If rsTmp.BOF = False Then strCode = strCode & ",'" & NVL(rsTmp("LIS��ϱ���")) & "'"

            lng�����Ŀid = 0
                        
            '�������Ŀid
            mstrSQL = "Select e.id " & _
                        "From   ������ĿĿ¼_LIS a," & _
                                "������ĿĿ¼ b," & _
                                "���鱨����Ŀ c," & _
                                "���鱨����Ŀ d," & _
                                "������ĿĿ¼ e " & _
                        "where  a.������Ŀid=b.id " & _
                                "and Nvl(b.�����Ŀ,0)=0 " & _
                                "and c.������Ŀid=b.id " & _
                                "and c.������Ŀid=d.������Ŀid " & _
                                "and d.������Ŀid=e.id " & _
                                "and e.�����Ŀ=1 and instr(','||a.LIS����||',','," & NVL(rs("������Ŀ����")) & ",')>0"
                        
            If rsTmp.State = adStateOpen Then rsTmp.Close
            rsTmp.Open mstrSQL, gcnOracle
            If rsTmp.BOF = False Then
                
                lng�����Ŀid = NVL(rsTmp("ID"))
                
                mstrSQL = "Select   C.ҽ��id As ����id " & _
                            "From   �����Ŀ�嵥 B," & _
                                    "�����Ŀҽ�� C, " & _
                                    "���鱨����Ŀ D " & _
                            "Where  B.������Ŀid= " & lng�����Ŀid & _
                                    " AND C.�嵥ID=B.ID  " & _
                                    " AND C.����id=" & lng����id & " " & _
                                    " AND B.�Ǽ�id=" & lng�Ǽ�id
                                    
    '            mstrSQL = "Select C.ҽ��id As ����id " & _
    '                        "From   (Select ������Ŀid From ������ĿĿ¼_LIS Where LIS���� In (" & strCode & ")) A, " & _
    '                                "�����Ŀ�嵥 B," & _
    '                                "�����Ŀҽ�� C " & _
    '                        "Where  A.������Ŀid=B.������Ŀid " & _
    '                                "AND C.�嵥ID=B.ID " & _
    '                                "AND C.����id=" & lng����id & " " & _
    '                                "AND B.�Ǽ�id=" & lng�Ǽ�id
                
                If rsTmp.State = adStateOpen Then rsTmp.Close
                rsTmp.Open mstrSQL, gcnOracle
                If rsTmp.BOF = False Then lng����id = NVL(rsTmp("����id"), 0)
                
                If lng����id > 0 Then
                    
                    '2.�Ҷ�Ӧ����Ŀ
                    strCode = "'" & NVL(rs("������Ŀ����")) & "'"
                    
                    mstrSQL = "Select C.ID,C.������ " & _
                                "From   ������ĿĿ¼_LIS A," & _
                                        "���鱨����Ŀ B," & _
                                        "����������Ŀ C  " & _
                                "Where C.ID=B.������Ŀid AND A.������Ŀid=B.������Ŀid AND A.LIS����=" & strCode
                    
                    If rsTmp.State = adStateOpen Then rsTmp.Close
                    rsTmp.Open mstrSQL, gcnOracle
                    
                    If rsTmp.BOF = False Then
                        mstrSQL = "ZL_ZLLIS_��д���("
                        
                        mstrSQL = mstrSQL & lng����id & ","
                        mstrSQL = mstrSQL & "'" & Trim(NVL(rsTmp("������"))) & "',"
                        mstrSQL = mstrSQL & NVL(rsTmp("ID"), 0) & ","
                        
                        If IsNull(rs("��ֵ�Խ��").Value) = False Then
                            mstrSQL = mstrSQL & "'" & Trim(NVL(rs("��ֵ�Խ��"))) & "',"
                            mstrSQL = mstrSQL & "1,"
                        Else
                            mstrSQL = mstrSQL & "'" & Trim(NVL(rs("�ַ��Խ��"))) & "',"
                            mstrSQL = mstrSQL & "0,"
                        End If
                        
                        mstrSQL = mstrSQL & "'" & Trim(NVL(rs("��λ").Value)) & "',"
                        
                        Select Case Trim(NVL(rs("�쳣���")))
                        Case "L", "l"
                            mstrSQL = mstrSQL & "'ƫ��',"
                        Case "H", "h"
                            mstrSQL = mstrSQL & "'ƫ��',"
                        Case Else
                            mstrSQL = mstrSQL & "'����',"
                        End Select
                        
                        mstrSQL = mstrSQL & "'" & Trim(NVL(rs("���뷶Χ"))) & "',"
                        mstrSQL = mstrSQL & "'" & Trim(NVL(rs("������"))) & "'"
                        
                        mstrSQL = mstrSQL & ")"
                        
                        gcnOracle.Execute mstrSQL, , adCmdStoredProc
                        
                    End If
                End If
            End If
            rs.MoveNext
        Loop
    End If
    
    gcnOracle.CommitTrans
    
    AcceptOneResult = True
    
    Exit Function
    
errHand:
    
    ShowSimpleMsg Err.Description
    gcnOracle.RollbackTrans
    
End Function

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    Dim strUser As String
    Dim strPsw As String
    Dim strSvr As String
    
    strUser = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ_LIS", "USER", "HISJQ")
    strSvr = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ_LIS", "SERVER", "")
    
    If frmLisLogin.ShowLogin(Me, strUser, strPsw, strSvr) Then
        
        If ConnectSQLServer(strSvr, "CliMis", strUser, strPsw) = False Then
            Unload Me
            Exit Sub
        End If
    Else
        Unload Me
        Exit Sub
    End If
    
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ_LIS", "USER", strUser
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ_LIS", "SERVER", strSvr
    
    Call ReadUnit
    
    If Not (lvw.SelectedItem Is Nothing) Then
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
    
End Sub

Private Sub Form_Load()
    Dim strVsf As String
    
    mblnStartUp = True
    
    strVsf = "��쵥��,900,1,1,1,;����,900,1,1,1,;�Ա�,600,1,1,1,;���֤��,1800,1,1,1,;�����,1200,1,1,1,;�Ǽ�id,0,1,1,1,;������λ,1200,1,1,1,"
    Call CreateVsf(vsf, strVsf)
'
'    With vsfItem
'        .Cols = 4
'
'        .TextMatrix(0, mCol.��Ŀ) = "��Ŀ"
'        .TextMatrix(0, mCol.���) = "���"
'        .TextMatrix(0, mCol.��־) = "��־"
'        .TextMatrix(0, mCol.�ο�) = "�ο�"
'
'        .ColWidth(mCol.��Ŀ) = 1500
'        .ColWidth(mCol.���) = 1200
'        .ColWidth(mCol.��־) = 600
'        .ColWidth(mCol.�ο�) = 1500
'    End With
    
    mstrKey = ""
    mlngTotal = Val(GetSetting("ZLSOFT", "����ȫ��\����ӿ�", "���ܼ��", "10"))
    If mlngTotal < 5 Then mlngTotal = 5
    If mlngTotal > 30 Then mlngTotal = 30
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With lvw
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
       
    With imgY
        .Top = lvw.Top
        .Height = lvw.Height
    End With
    
    With vsf
        .Left = imgY.Left + imgY.Width
        .Top = lvw.Top + 30
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
    
'    With vsfItem
'        .Left = vsf.Left
'        .Top = vsf.Top + vsf.Height + 45
'        .Width = vsf.Width
'    End With
    
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If mstrKey <> Item.Key Then
        
        mstrKey = Item.Key
        Call ReadPerson(Val(Mid(Item.Key, 2)))
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
        
    End If
    
    
End Sub

Private Sub mnuEditAcceptPerson_Click()
    Dim blnSvr As Boolean
    
    If Val(vsf.RowData(vsf.Row)) <= 0 Then Exit Sub
    
    If MsgBox("ȷʵ��Ҫ���ܴ��˵ļ���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    blnSvr = tmr.Enabled
    tmr.Enabled = False
    
    frmWait.OpenWait Me, "���ܼ�������", True
    frmWait.WaitInfo = "���ڽ��ܡ�" & vsf.TextMatrix(vsf.Row, mCol.����) & "����������"
    
    If AcceptOneResult(Val(vsf.TextMatrix(vsf.Row, mCol.�Ǽ�id)), Val(vsf.RowData(vsf.Row))) Then
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    End If
    
    frmWait.CloseWait
    tmr.Enabled = blnSvr
    
End Sub

Private Sub mnuEditAutoAccept_Click()
    '
    mnuEditAutoAccept.Checked = Not mnuEditAutoAccept.Checked
    
    If mnuEditAutoAccept.Checked Then
        tbrThis.Buttons("����").Value = tbrPressed
        tmr.Enabled = True
    Else
        tbrThis.Buttons("����").Value = tbrUnpressed
        tmr.Enabled = False
    End If
    
End Sub

Private Sub mnuEditResetPerson_Click()
    
    Dim rs As New ADODB.Recordset
    
    If Val(vsf.RowData(vsf.Row)) = 0 Then Exit Sub
    
    On Error GoTo errHand
    
    If MsgBox("ȷʵ��Ҫ��������ѽ��ܵ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If ClearOneResult(Val(vsf.TextMatrix(vsf.Row, mCol.�Ǽ�id)), Val(vsf.RowData(vsf.Row))) Then
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    End If
    
    Exit Sub
    
errHand:
    ShowSimpleMsg Err.Description
    
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileParam_Click()
    
    If frmTaskAcceptParam.ShowParam(Me) Then
        mlngTotal = Val(GetSetting("ZLSOFT", "����ȫ��\����ӿ�", "���ܼ��", "10"))
        If mlngTotal < 5 Then mlngTotal = 5
        If mlngTotal > 30 Then mlngTotal = 30
    End If
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuHelpTopic_Click()
    Call ShowHelp(Me.hWnd, Me.Name)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "����"
        Call mnuEditAutoAccept_Click
    Case "ָ��"
        Call mnuEditAcceptPerson_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tmr_Timer()
    
    mlngCount = mlngCount + 1
    
    If mlngCount >= mlngTotal Then
        mlngCount = 0
        tmr.Enabled = False
        If AcceptResult Then
            Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
        End If
        tmr.Enabled = True
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If OldRow <> NewRow Then
        Call ReadItems(Val(vsf.TextMatrix(NewRow, mCol.�Ǽ�id)), Val(vsf.RowData(NewRow)))
    End If
    
End Sub

