VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   7035
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   8295
      Begin VSFlex8Ctl.VSFlexGrid vsfSub 
         Height          =   5445
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6915
         _cx             =   12197
         _cy             =   9604
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
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
         GridColor       =   16777215
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
   End
   Begin VB.Frame fraType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ҩƷ����"
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2175
      Begin MSComctlLib.TreeView tvwType 
         Height          =   6495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   11456
         _Version        =   393217
         Indentation     =   354
         LabelEdit       =   1
         LineStyle       =   1
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ҩƷ����"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_conn As ADODB.Connection
Private m_mease As New MAES
Private m_base64 As New DecodeBase64
Dim m_str As String
Private Const WM_USER = &H400
Private Const EM_EXSETSEL = (WM_USER + 55)
Private Const EM_SETSEL = &HB1
Private Const EM_GETSEL = &HB0
Private Const EM_GETPARAFORMAT = (WM_USER + 61)
Private Const EM_SETPARAFORMAT = (WM_USER + 71)
Private Const EM_GETSELTEXT = (WM_USER + 62)
Private Const EM_SETTYPOGRAPHYOPTIONS = (WM_USER + 202)
Private Const EM_GETTYPOGRAPHYOPTIONS = (WM_USER + 203)
Private Const TO_ADVANCEDTYPOGRAPHY = &H1
Private Const TO_SIMPLELINEBREAK = &H2&
Private Const PFM_ALIGNMENT = &H8
Private Const PFM_TABSTOPS = &H10
Private Const PFM_STYLE = &H400
Private Const PFA_LEFT = 1
Private Const PFA_RIGHT = 2
Private Const PFA_CENTER = 3
Private Const PFA_JUSTIFY = &H4
Private Const PS_SOLID = 0
Private Const PFA_FULL_GLYPHS = 7
Private Const mZERO = &H0&
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
 
Private Type PARAFORMAT2
    cbsize As Integer
    dwpad As Integer
    dwMask As Long
    wNumbering As Integer
    wReserved As Integer
    dxStartIndent As Long
    dxRightIndent As Long
    dxOffset As Long
    wAlignment As Integer
    cTabCount As Integer
    lTabstops(0 To 31&) As Long
    dySpaceBefore As Long
    dySpaceAfter As Long
    dyLineSpacing As Long
    sStyle As Integer
    bLineSpacingRule As Byte
    bOutlineLevel As Byte
    wShadingWeight As Integer
    wShadingStyle As Integer
    wNumberingStart As Integer
    wNumberingStyle As Integer
    wNumberingTab As Integer
    wBorderSpace As Integer
    wBorderWidth As Integer
    wBorders As Integer
 
End Type

Public Enum ERECParagraphLineSpacingConstants
ercLineSpacingSingle = 0
ercLineSpacingOneAndAHalf = 1
ercLineSpacingDouble = 2
ercLineSpacingTwips = 3
ercLineSpacingTwipsAnyMinimum = 4
ercLineSpacingTwentiethLine = 5
End Enum

Private Const PFM_SPACEBEFORE = &H40&
Private Const PFM_SPACEAFTER = &H80&
Private Const PFM_LINESPACING = &H100&
Private Const PFM_BORDER = &H800&                   ' /* (*)  */
Private Const PFM_SHADING = &H1000&                 ' /* (*)  */
Private Const PFM_NUMBERINGSTYLE = &H2000&          ' /* (*)  */
Private Const PFM_NUMBERINGTAB = &H4000&            ' /* (*)  */
Private Const PFM_NUMBERINGSTART = &H8000&         ' /* (*)  */


Public Sub InitDataByString(connectionString As String, str As String)
    Set m_conn = New ADODB.Connection
    m_conn.connectionString = connectionString
    m_str = UCase(Algorithm(m_base64.EncodeBase64String(str), 32))
    m_conn.Open
    m_conn.CursorLocation = adUseClient
End Sub

Public Sub InitDataByADODB(conn As ADODB.Connection, str As String)
    Set m_conn = conn
    m_str = UCase(Algorithm(m_base64.EncodeBase64String(str), 32))

End Sub

Public Function ReadLob(ByVal strKey As String, ByVal strCol As String) As String
    Const conChunkSize As Integer = 10240
    
    Dim rsLob As ADODB.Recordset
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim strSQL As String
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select Zl_Drugexplain_Readlob(?,?,?) as Ƭ�� From Dual"

    'CLOB
    lngCount = 0
    strFile = ""
    Do
        Dim cmdData As New ADODB.Command
        
        Set cmdData.ActiveConnection = m_conn
        cmdData.Parameters.Append cmdData.CreateParameter("PAR1", adVarChar, adParamInput, LenB(StrConv(strKey, vbFromUnicode)), strKey)
        cmdData.Parameters.Append cmdData.CreateParameter("PAR2", adVarChar, adParamInput, LenB(StrConv(strCol, vbFromUnicode)), strCol)
        cmdData.Parameters.Append cmdData.CreateParameter("PAR3", adVarNumeric, adParamInput, 30, lngCount)
        
        cmdData.CommandText = strSQL
        On Error Resume Next
        Set rsLob = cmdData.Execute
        If Err.Number <> 0 Then Err.Clear: Exit Do
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).value) Then Exit Do
        strText = rsLob.Fields(0).value
        strFile = strFile & strText
        lngCount = lngCount + 1
    Loop

    ReadLob = strFile
    Exit Function
Errhand:
    Err.Clear
End Function

Public Sub LoadContent(drugsid As String)
    Dim myrs As New ADODB.Recordset
    Dim sqlstr As String
    Dim lngTmp As Long
    Dim cmdData As New ADODB.Command
    
    sqlstr = "select t.ͨ������,t.��Ʒ��,t.����ƴ��,t.Ӣ������,t.ҩ����,t.ҩ�����,t.������ҵ,t.��׼�ĺ�,'' ��ѧ����,''  ��״,'' ҩ����,'' ҩ������ѧ,'' ��Ӧ֢,'' �÷�����,'' ������Ӧ,'' ����֢,'' ע������,'' �и���ҩ,'' ��ͯ��ҩ,'' ��������ҩ,'' �໥����,'' ҩ�����,'' �������� from ҩƷ˵���� t where id= ?"
    On Error GoTo errH
    vsfSub.Rows = 0
    Set cmdData.ActiveConnection = m_conn
    cmdData.Parameters.Append cmdData.CreateParameter("PAR1", adVarChar, adParamInput, LenB(StrConv(drugsid, vbFromUnicode)), drugsid)
    cmdData.CommandText = sqlstr
    Set myrs = cmdData.Execute
    Set myrs = CopyNewRec(myrs)

    myrs!��ѧ���� = ReadLob(drugsid, "��ѧ����")
    myrs!��״ = ReadLob(drugsid, "��״")
    myrs!ҩ���� = ReadLob(drugsid, "ҩ����")
    myrs!ҩ������ѧ = ReadLob(drugsid, "ҩ������ѧ")
    myrs!��Ӧ֢ = ReadLob(drugsid, "��Ӧ֢")
    myrs!�÷����� = ReadLob(drugsid, "�÷�����")
    myrs!������Ӧ = ReadLob(drugsid, "������Ӧ")
    myrs!����֢ = ReadLob(drugsid, "����֢")
    myrs!ע������ = ReadLob(drugsid, "ע������")
    myrs!�и���ҩ = ReadLob(drugsid, "�и���ҩ")
    myrs!��ͯ��ҩ = ReadLob(drugsid, "��ͯ��ҩ")
    myrs!��������ҩ = ReadLob(drugsid, "��������ҩ")
    myrs!�໥���� = ReadLob(drugsid, "�໥����")
    myrs!ҩ����� = ReadLob(drugsid, "ҩ�����")
    myrs!�������� = ReadLob(drugsid, "��������")
    
    If myrs.EOF = False Then
      Dim index As Integer
      Dim name As String
      For I = 0 To myrs.Fields.Count - 1
          Set myrd = myrs(I)
          name = "��" & myrd.name & "��"
          vsfSub.Rows = vsfSub.Rows + 1
          vsfSub.TextMatrix(vsfSub.Rows - 1, 0) = name
          vsfSub.Cell(flexcpForeColor, vsfSub.Rows - 1, 0, vsfSub.Rows - 1, 0) = vbBlue
          If IsNull(myrd.value) = False Then
              If (myrd.Type = adInteger Or myrd.Type = adNumeric) Then
                  vsfSub.Rows = vsfSub.Rows + 1
                  vsfSub.TextMatrix(vsfSub.Rows - 1, 0) = "   " & Replace(myrd.value, vbCrLf, "")
              Else
                  vsfSub.Rows = vsfSub.Rows + 1
                  vsfSub.TextMatrix(vsfSub.Rows - 1, 0) = "   " & Replace(m_mease.DecryptStr(myrd.value, m_str, Bit256, Bit128, False), vbCrLf, "")
              End If
              If myrd.name = "ͨ������" Then
                  If (myrd.Type = adInteger Or myrd.Type = adNumeric) Then
                      lblName.Caption = Replace(myrd.value, vbCrLf, "")
                  Else
                      lblName.Caption = Replace(m_mease.DecryptStr(myrd.value, m_str, Bit256, Bit128, False), vbCrLf, "")
                  End If
              End If
          End If
      Next
    End If
    Exit Sub
errH:
    Err.Clear
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'���Ƽ�¼��
'������strFields=��Ҫ���Ƶļ�¼�����ֶε���˳����ֶ�����ɵ��ַ���
'          �磺1 ����1,3 ����2,7 ����3...��ʾ���Ƽ�¼���ĵ�1,3,7..�ֶ���ɼ�¼��������
'              ID ����1,���� ����2,....��ʾ���Ƽ�¼����ID,����...�ֶ���ɼ�¼������
'              ����*Ϊ�µļ�¼��������
'              �������ͻ�����׳���������ͬ�����⣬��ע��
'           arrAppFields=׷�ӵ��ֶ���Ϣ������,����,����,Ĭ��ֵ,û��Ĭ��ֵ��Empty,û��ָ�����ȴ�Empty
'      blnOnlyStructure=�Ƿ�ֻ���ƽṹ
'�ڳ����У��������漰���໥���ݼ�¼������ʹ��ADO��Clone���Ʋ����ļ�¼����������һ����¼�������ݷ����仯��ʱ�����и�������������ͬ�ı仯��ͨ��ָ�޸Ļ�ɾ����������������ϣ����Щ��¼���໥�䱣�ֶ���
  
    Dim rsClone As ADODB.Recordset
    Dim rsTarget As ADODB.Recordset
    Dim intFields As Integer
    Dim arrFieldsName As Variant, strFieldName As String, strFieldNameAlias As String
    Dim arrTmp As Variant
    Dim I As Long
    
    On Error GoTo errH
    If Not rsSource Is Nothing Then
        Set rsClone = rsSource.Clone
        rsClone.Filter = rsSource.Filter
    End If
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        '������¼���ṹ
        If Not rsClone Is Nothing Then
            If strFields = "" Then '��¼��ȫ����ģʽ
                arrFieldsName = Array()
                If rsClone.Fields.Count > 0 Then
                    ReDim arrFieldsName(rsClone.Fields.Count - 1)
                Else
                    arrFieldsName = Array()
                End If
                For intFields = 0 To rsClone.Fields.Count - 1
                    arrFieldsName(intFields) = rsClone.Fields(intFields).name & ""
                    .Fields.Append rsClone.Fields(intFields).name, IIf(rsClone.Fields(intFields).Type = adNumeric, adDouble, rsClone.Fields(intFields).Type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
                Next
            Else '��¼�����ָ���ģʽ
                If rsClone.Fields.Count > 0 Then
                    arrFieldsName = Split(strFields, ",")
                    For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                        '�а�������
                        arrTmp = Split(arrFieldsName(intFields) & " ", " ")
                        strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                        If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).name & ""
                        '��ȡ�ֶ�ԭ������������
                        arrFieldsName(intFields) = strFieldName
                        '����ֶ�,�������ڱ������������е�����Ϊ����
                        .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:��ʾ����
                    Next
                End If
            End If
        End If
        '׷���ֶ����
        If TypeName(arrAppFields) = "Variant()" Then
            For I = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(I + 2) = Empty Then
                    If arrAppFields(I + 3) = Empty Then
                        .Fields.Append arrAppFields(I), arrAppFields(I + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(I), arrAppFields(I + 1), , adFldIsNullable, arrAppFields(I + 3)
                    End If
                Else
                    If arrAppFields(I + 3) = Empty Then
                        .Fields.Append arrAppFields(I), arrAppFields(I + 1), arrAppFields(I + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(I), arrAppFields(I + 1), arrAppFields(I + 2), adFldIsNullable, arrAppFields(I + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '��������
        If Not blnOnlyStructure And Not rsClone Is Nothing Then
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '�¼�¼�����а�˳����ӣ���˿�������
                    .Fields(intFields).value = rsClone.Fields(arrFieldsName(intFields)).value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
    Exit Function
errH:
    Err.Clear
End Function

Private Sub Form_Load()
    '��ʼ��ҩƷ���Է���ؼ�
    Call SetMainDirectory

    vsfSub.AllowSelection = False 'ֻ������ѡ��
    vsfSub.SelectionMode = flexSelectionByRow '����ѡ����ʽ
End Sub

Public Sub SetMainDirectory()
'����:����ҳ�����õ���Ŀ¼
    Dim myNod As Node
    Dim strTmp As String
    Dim strTitle() As String
    Dim I As Long, J As Long
    
    On Error GoTo errH
    strTmp = "����,��ѧ����,��״,ҩ����,ҩ������ѧ,��Ӧ֢,�÷�����,������Ӧ,����֢,ע������,�и���ҩ,��ͯ��ҩ,��������ҩ,�໥����,ҩ�����,��������"
    strTitle = Split(strTmp, ",")
    tvwType.Nodes.Clear
    tvwType.LineStyle = tvwRootLines
    tvwType.Indentation = 200
    For I = LBound(strTitle) To UBound(strTitle)
        Set myNod = tvwType.Nodes.Add(, , "key-" & I, I + 1 & "." & strTitle(I))
        myNod.Tag = strTitle(I)
        myNod.Expanded = True
    Next
    Exit Sub
errH:
    Err.Clear
End Sub


Private Sub Form_Resize()
On Error GoTo errH
    fraName.Move 100, 100, Me.Width - 200, 855
    If Me.Height - (fraName.Top + fraName.Height + 100) - 100 >= 0 Then
        fraType.Move 100, fraName.Top + fraName.Height + 100, 2175, Me.Height - (fraName.Top + fraName.Height + 100) - 100
        tvwType.Move 120, 240, 1935, IIf(fraType.Height - 340 > 0, fraType.Height - 340, 0)
    End If
    If Me.Width - (fraType.Left + fraType.Width + 100) - 100 >= 0 And Me.Height - (fraName.Top + fraName.Height + 100) - 100 >= 0 Then
        fraInfo.Move fraType.Left + fraType.Width + 100, fraType.Top, Me.Width - (fraType.Left + fraType.Width + 100) - 100, Me.Height - (fraName.Top + fraName.Height + 100) - 100
        vsfSub.Move 100, 300, IIf(fraInfo.Width - 200 > 0, fraInfo.Width - 200, 0), IIf(fraInfo.Height - 400 > 0, fraInfo.Height - 400, 0)
        vsfSub.ColWidthMax = IIf(vsfSub.Width - 200 > 0, vsfSub.Width - 200, 0)
        vsfSub.ColWidthMin = IIf(vsfSub.Width - 200 > 0, vsfSub.Width - 200, 0)
        '��ʼ�����
        vsfSub.WordWrap = True    '���ֻ���
        vsfSub.AutoSizeMode = flexAutoSizeRowHeight '�Զ�����
        vsfSub.AutoSize 0
    End If
    Exit Sub
errH:
    Err.Clear
End Sub


Private Sub tvwType_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strType As String
On Error GoTo errH
    If Node.Tag <> "" And vsfSub.Rows <> 0 Then
        If Node.Tag <> "����" Then
            strType = "��" & Node.Tag & "��"
            vsfSub.Row = vsfSub.FindRow(strType, , 0)
        Else
            vsfSub.Row = 0
        End If
        vsfSub.ShowCell vsfSub.Row, 0
        vsfSub.ShowCell vsfSub.Row + 1, 0
    End If
    Exit Sub
errH:
    Err.Clear
End Sub

Private Sub vsfSub_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub vsfSub_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub vsfSub_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

