VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.UserControl usrVerifyReport 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000E&
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   ScaleHeight     =   3225
   ScaleWidth      =   5700
   Begin VB.PictureBox picComment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   1980
      TabIndex        =   2
      Top             =   2085
      Width           =   2010
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   465
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   -15
         Width           =   2160
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���鱸ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   390
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf2 
      Height          =   1335
      Left            =   630
      TabIndex        =   0
      Top             =   540
      Width           =   2355
      _cx             =   4154
      _cy             =   2355
      Appearance      =   0
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
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483645
      GridColorFixed  =   -2147483645
      TreeColor       =   -2147483639
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   240
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin zl9CISCore.VsfGrid vsf 
      Height          =   1530
      Left            =   1470
      TabIndex        =   1
      Top             =   0
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   2699
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000003&
      Height          =   435
      Left            =   0
      Top             =   150
      Width           =   585
   End
End
Attribute VB_Name = "usrVerifyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private rsTmp As New ADODB.Recordset

Private mlng����id As Long                      '��紫��
Private mlngҽ��id As Long                      '��紫��
Private mstr�Ա� As String

Private mblnMode As Boolean 'Ϊ���Ǳ�ʾ���û����еı༭����ʱ�Ÿ�ֵ

Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String
Private mblnMoved As Boolean '�����Ƿ���ת��

Private mblnModify As Boolean
Private mblnCommon As Boolean

Private mstrSQL As String
Private mlngLoop As Long
Private mblnLoaded As Boolean

Private Enum mCol
    ������Ŀ = 1
    ������
    �����־
    �������
    ���㹫ʽ
    ��λ
    ����ο�
    
    ϸ�� = 0
    ����
    ������
    ���
    ����
    ��������
End Enum

Private Enum COLOR
    
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ��ɫ = &H40C0&
    ��������ɫ = &H40C0&
    ����ǰ��ɫ = &H8000000E
    ���걳��ɫ = &H80C0FF
    �ͱ걳��ɫ = &H80FFFF
    ����ǰ��ɫ = &H80000012
    Ĭ��ǰ��ɫ = &H80000008
End Enum

Private mobjParentObject As Object

Public Property Set ParentObject(vData As Object)
    Set mobjParentObject = vData
End Property

Public Property Get ParentObject() As Object
    Set ParentObject = mobjParentObject
End Property

Private Property Let Modified(vData As Boolean)
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    mobjParentObject.Modified = vData
    
End Property

Private Property Get Modified() As Boolean
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    Modified = mobjParentObject.Modified
    
End Property

'-------------------------------------------------------------------------------------------------------------------
Private Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:������ؼ��Ŀ���
    '����:objVsf Ҫ�����еı��ؼ�����
    '����:���ɹ�����True,���򷵻� False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.�������е���
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.���¼�����Ҫ������
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.���¼�����Ҫ�ĺ���
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
End Function
Private Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '������
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '��С��
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Private Sub ApplyResultColor(vsf As Object, ByVal lngRow As Long, ByVal lngCol As Long, ByVal bytMode As Byte)
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    Dim lngColor As Long, lngForeColor As Long
    
    Select Case bytMode
        Case 0, 1
            lngColor = &H80000005
            lngForeColor = COLOR.Ĭ��ǰ��ɫ
        Case 5 '�쳣�͡���
            lngColor = COLOR.��������ɫ
            lngForeColor = COLOR.����ǰ��ɫ
        Case 2
            lngColor = COLOR.�ͱ걳��ɫ
            lngForeColor = COLOR.����ǰ��ɫ
        Case Else
            lngColor = COLOR.���걳��ɫ
            lngForeColor = COLOR.����ǰ��ɫ
    End Select
    
    vsf.Cell(flexcpBackColor, lngRow, lngCol, lngRow, lngCol) = lngColor
    vsf.Cell(flexcpForeColor, lngRow, lngCol, lngRow, lngCol) = lngForeColor
End Sub

Private Function CalcDefaultFlag(ByVal strValue As String, ByVal strReference As String, Optional ByVal bytMode As Byte = 1, _
    Optional ByVal strAlarmLow As String, Optional ByVal strAlarmHigh As String) As String
    
    '--------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '--------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    If Len(Trim(strValue)) = 0 Then CalcDefaultFlag = "": Exit Function
    
    CalcDefaultFlag = ""
    
    If InStr(strReference, vbCrLf) > 0 Then strReference = Mid(strReference, 1, InStr(strReference, vbCrLf) - 1)
    If Trim(strReference) = "" Then Exit Function
                
    If bytMode = 2 Or bytMode = 3 Then '���ԡ��붨��
        If bytMode = 2 Or InStr(strReference, "��") = 0 Or Trim(strValue) Like "*��*" Or Trim(strValue) Like "*+*" Or _
            Trim(strValue) Like "*��*" Or Trim(strValue) Like "*��*" Or Trim(strValue) Like "*-*" Then
            '���Ի��޷�Χ�ο��İ붨��
            If (Len(Trim(strReference)) > 0 And (Trim(strReference) Like (Trim(strValue) & "*") Or Trim(strReference) Like ("*" & Trim(strValue)))) Or _
                (Not (Trim(strValue) Like "*��*" Or Trim(strValue) Like "*+*" Or Trim(strValue) Like "*��*")) Then
                CalcDefaultFlag = ""
            Else
                CalcDefaultFlag = "�쳣"
            End If
            Exit Function
        Else
            '��ȡ�붨��ֵ
            For i = 1 To Len(Trim(strValue))
                If InStr("01234567890.", Mid(strValue, i, 1)) > 0 Then Exit For
            Next
            If i > Len(Trim(strValue)) Then Exit Function
            strValue = Val(Mid(strValue, i))
        End If
    End If
    '�ߵ��ж�
    If Len(Trim(strAlarmLow)) > 0 And Val(strAlarmLow) <> 0 Then
        If Val(strValue) < Val(strAlarmLow) Then
            CalcDefaultFlag = "����"
            Exit Function
        End If
    End If
    If Len(Trim(strAlarmHigh)) > 0 And Val(strAlarmHigh) <> 0 Then
        If Val(strValue) > Val(strAlarmHigh) Then
            CalcDefaultFlag = "����"
            Exit Function
        End If
    End If
    If InStr(strReference, "��") > 0 Then
        
        '���С�ڲο���ֵ
        If Val(strValue) < Val(Mid(strReference, 1, InStr(strReference, "��") - 1)) And _
            Len(Trim(Mid(strReference, 1, InStr(strReference, "��") - 1))) > 0 Then
            CalcDefaultFlag = "��"
        End If
        
        '������ڲο���ֵ
        If Val(strValue) > Val(Mid(strReference, InStr(strReference, "��") + 1)) And _
            Len(Trim(Mid(strReference, InStr(strReference, "��") + 1))) > 0 Then
            CalcDefaultFlag = "��"
        End If
            
    End If
End Function

Private Function CalcExpress(ByVal vsf As Object, ByVal strExPress As String) As Single
    
    '--------------------------------------------------------------------------------------------------------
    '����:�ڱ���м���ĳһ���ʽ�Ľ��
    '����:vsf           ������ݵı��
    '     strExpress    Ҫ����ı��ʽ
    '����:������ֵ
    '--------------------------------------------------------------------------------------------------------
    
    Dim strTmpPress As String
    Dim rs As New ADODB.Recordset
    
    Dim lngTmpID As Long
    Dim lngLeftPos As Long
    Dim lngRightPos As Long
    Dim lngLoop As Long
    Dim sglValue As Single
    
    CalcExpress = 0
    
    strTmpPress = strExPress
    If strTmpPress <> "" Then
        
        lngLeftPos = InStr(strTmpPress, "[")
        lngRightPos = InStr(strTmpPress, "]")
        
        Do While lngLeftPos > 0
        
            lngTmpID = Val(Mid(strTmpPress, lngLeftPos + 1, lngRightPos - lngLeftPos - 1))
            
            '�ж�lngTmpID�Ƿ�Ҳ�Ǽ�����Ŀ
            For lngLoop = 1 To vsf.Rows - 1
                If Val(vsf.RowData(lngLoop)) = lngTmpID Then
                    If Trim(vsf.TextMatrix(lngLoop, mCol.���㹫ʽ)) <> "" Then
                        '�Ǽ�����Ŀ,�ȼ�����˽��
                        sglValue = CalcExpress(vsf, Trim(vsf.TextMatrix(lngLoop, mCol.���㹫ʽ)))
                    Else
                        '���Ǽ�����Ŀ,ֱ��ȡ�˽��
                        sglValue = Val(vsf.TextMatrix(lngLoop, mCol.������))
                    End If
                    
                    Exit For
                    
                End If
            Next
            
            '�ڵ�ǰ�����û�д˼�����Ŀ,��Ϊ���Ϊ��
            If lngLoop = vsf.Rows Then sglValue = 0
                                        
            '�Խ��������ʽ�еļ�������
            strTmpPress = Mid(strTmpPress, 1, lngLeftPos - 1) & sglValue & Mid(strTmpPress, lngRightPos + 1)
            
            '����һ���������ӵ�λ��
            lngLeftPos = InStr(strTmpPress, "[")
            lngRightPos = InStr(strTmpPress, "]")
        Loop
                
        '������ʽ�Ľ��
        On Error Resume Next
        Call OpenRecord(rs, "SELECT " & strTmpPress & " AS ��� FROM DUAL", "������")
        If rs.BOF = False Then CalcExpress = zlCommFun.NVL(rs("���"), 0)
        On Error GoTo 0
        
    End If
    
End Function

Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub

Private Function ShowOpenList(Optional strText As String, Optional blnWhere As Boolean = False) As Byte
    '-----------------------------------------------------------------------------------------
    '����:���б�ṹ�����Ƽ���걾����
    '����:������2;�ɹ�����1;ȡ������0
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strLvw = "����,900,0,1;ȡֵ,1800,0,0;�����־,900,0,0"

    ShowOpenList = 2
    
    strSQL = "SELECT ROWNUM AS ID,����,ȡֵ,DECODE(�����־,1,'',2,'��',3,'��',4,'�쳣',5,'����',6,'����','') AS �����־ FROM ������Ŀȡֵ A WHERE ��Ŀid=[1]"
    
    If blnWhere Then
        strSQL = strSQL & " AND (A.���� Like [2] OR A.ȡֵ Like [2])"
    End If
        
    Set rs = OpenSQLRecord(strSQL, "������", CLng(Val(vsf.RowData(vsf.Row))), "%" & strText & "%")
    
    If rs.BOF Then
        
        ShowOpenList = 0
        
        Exit Function
    End If
    
    If rs.RecordCount = 1 And blnWhere Then GoTo Over
        
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 3600, 4500, "����ȡֵѡ��", "����±���ѡ��һ��ȡֵ") Then
        GoTo Over
    End If
    
    Exit Function
    
Over:
    vsf.EditText = zlCommFun.NVL(rs("ȡֵ").Value)
    vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("ȡֵ").Value)
    vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("ȡֵ").Value)
    vsf.TextMatrix(vsf.Row, mCol.�����־) = zlCommFun.NVL(rs("�����־").Value)
    
    Call ApplyResultColor(vsf, vsf.Row, mCol.������, _
            Decode(vsf.TextMatrix(vsf.Row, mCol.�����־), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

'��������������
Public Property Get DispMode() As Boolean
    '�Ƿ�Ϊ��ʾģʽ
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    mDispMode = New_DispMode
    ShowUsrControl mlngҽ��id, Not mDispMode
    PropertyChanged "DispMode"
    
    If mDispMode Then
        vsf.Body.Editable = flexEDNone
    End If
    
End Property

Public Property Get ID���˲���() As Long
    '���ز��˲���ID
    
    ID���˲��� = mlng����id
End Property

Public Property Let ID���˲���(ByVal New_ID���˲��� As Long)
    '���ò��˲���ID,�����ò����ǲ��Ǵ���
    
    mlng����id = New_ID���˲���
    ShowUsrControl mlngҽ��id, Not mDispMode
    
End Property

Public Sub SetDiagItem(ByVal New_ҽ��ID As Long, ByVal New_���ͺ�)
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    mlngҽ��id = New_ҽ��ID
    strSQL = "SELECT DECODE(���id,NULL,ID,���id) AS ID FROM ����ҽ����¼ WHERE ID=[1]"
    
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rs = OpenSQLRecord(strSQL, "���鱨��ר��ֽ", mlngҽ��id)
    If rs.BOF = False Then mlngҽ��id = rs("ID").Value
        
        
    mblnModify = True
    mblnCommon = True       '�Ƿ�Ϊ��ͨ������Ŀ
    
    strSQL = "SELECT DISTINCT A.ִ��״̬,D.��Ŀ��� FROM ����ҽ������ A,����ҽ����¼ B,���鱨����Ŀ C,������Ŀ D WHERE B.������Ŀid=C.������Ŀid AND C.������Ŀid=D.������Ŀid AND A.ҽ��id=B.ID AND B.���id=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rs = OpenSQLRecord(strSQL, "���鱨��ר��ֽ", mlngҽ��id)
    If rs.BOF = False Then
        mblnModify = (rs("ִ��״̬").Value <> 1)
        mblnCommon = (rs("��Ŀ���").Value <> 2)
    End If
                
    vsf.Visible = False
    vsf2.Visible = False
    
    If mblnCommon Then
        With vsf
            .Cols = 0
            .NewColumn "", 255, 4
            .NewColumn "��Ŀ", 2100, 1
            .NewColumn "���", 900, 1, , 1
            .NewColumn "��־", 990, 1, " ", 1
            .NewColumn "����", 0, 1
            .NewColumn "��ʽ", 0, 1
            .NewColumn "��λ", 0, 1
            .NewColumn "�ο�", 1500, 1
            
            .ExtendLastCol = True
            .Body.Appearance = flexFlat
            '.AppearanceFlat = True
            .Body.BorderStyle = flexBorderNone
            .Body.BackColorFixed = .Body.BackColor
            .FixedCols = 1
            
            .Cell(flexcpFontBold, 0, 0, 0, vsf.Cols - 1) = True
            .Visible = True
        End With
        
        If mblnModify = False Then
            vsf.EditMode(mCol.������) = 0
            vsf.EditMode(mCol.�����־) = 0
            vsf.ComboList(mCol.������) = ""
            vsf.ComboList(mCol.�����־) = ""
        End If
        
    Else
        With vsf2
            .Cols = 6
            .FixedRows = 2
            .FixedCols = 0
            
            .MergeRow(0) = True
            .MergeCol(0) = True
            .MergeCol(1) = True
            .MergeCol(5) = True
            
            .TextMatrix(0, 0) = "ϸ��"
            .TextMatrix(1, 0) = "ϸ��"
            
            .TextMatrix(0, 1) = "����"
            .TextMatrix(1, 1) = "����"
            
            .Cell(flexcpText, 0, 2, 0, 4) = "ҩ������"
            .TextMatrix(1, 2) = "������"
            .TextMatrix(1, 3) = "���"
            .TextMatrix(1, 4) = "����"
            
            .TextMatrix(0, 5) = "��������"
            .TextMatrix(1, 5) = "��������"
            
            .ColWidth(0) = 1800
            .ColWidth(1) = 510
            .ColWidth(2) = 1800
            .ColWidth(3) = 1200
            .ColWidth(4) = 810
            .ColWidth(5) = 30
            
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .ColAlignment(3) = 1
            .ColAlignment(4) = 1
            .ColAlignment(5) = 1
            
            .Cell(flexcpAlignment, 0, 0, 1, vsf2.Cols - 1) = 4
            
            .ExtendLastCol = True
            .Visible = True
            .Cell(flexcpFontBold, 0, 0, 1, vsf2.Cols - 1) = True
            Call AppendRows(vsf2, lnX, lnY)
            
            mblnModify = False
        End With
    End If
End Sub

Public Property Get Getҽ��id() As Long
        
    Getҽ��id = mlngҽ��id
        
End Property

Private Sub SetErr(lngErrNum As Long, strErr As String)
    '���ô��������������
    '���lngErrNum=-1 ��ʾ �ؼ��Լ�����Ĵ���
    mReturnErrnumber = lngErrNum
    mReturnErrDescription = strErr
End Sub

Public Property Get ReturnErrNumber() As Long
    '�������һ�εĴ����
    ReturnErrNumber = mReturnErrnumber
End Property

Public Property Get ReturnErrDescription() As String
    '�������һ�δ��������ַ���
    ReturnErrDescription = mReturnErrDescription
End Property

'-------------------------------------------------------------------------------------------------------------------
Private Function CheckStrValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByRef strError As String) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
        
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        strError = "���������ݺ��зǷ��ַ���"
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            strError = "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��"
            Exit Function
        End If
    End If
    
    CheckStrValid = True
End Function

'------------------------------------------------------------------------------------------------------------

Private Sub ShowUsrControl(lngKey As Long, Optional ByVal blnEditMode As Boolean = False)
    '------------------------------------------------------------------------------------------------------------
    '���ܣ��ⲿ������ʾ������Ҫ�Ĺ���
    '------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
        
    mDispMode = Not blnEditMode
    
    '���߼�Ӧ�ȳ�ʼ�ؼ�
    Call InitData
    
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Sub

    Call ReadData
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub SetgcnOracle()
    '-------------------------------------------------------------------------------------------------
    '�ӿ�
    '-------------------------------------------------------------------------------------------------
    
    Call InitCommon(gcnOracle)
End Sub

Private Sub InitData()
    '��ʼ������
    
    Dim strTmp As String
    
    On Error GoTo ErrHandle
        
    If Not gcnOracle Is Nothing Then
        If Not gcnOracle.State <> adStateOpen Then
            If Ambient.UserMode = True Then

            End If
        End If
    
    End If
    
    
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then
        SetErr Err.Number, Err.Description
        Exit Sub
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ��������ݿ��������
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strDec As String
    
    On Error GoTo ErrHandle
    
    mblnLoaded = True
    
    '���Ӽ��
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Function
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Function
                                        
    txt.Text = ""
    
    If mblnCommon Then
        '�������
        vsf.Rows = 2
        vsf.RowData(1) = 0
        vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        
        '��ȡ��װ������
    
    '    ������Ŀ = 0
    '    ������
    '    �����־
    '    ����ο�
    '    �������
    '    ���㹫ʽ
        
        mstrSQL = "SELECT ������id,�������� FROM ���˲��������� WHERE ����id=[1]"
        
        If mlng����id > 0 Then
            mstrSQL = "SELECT DISTINCT F.������id AS ID," & _
                               "G.������ AS ������Ŀ," & _
                               "F.�������� AS ������," & _
                               "zlGetReference(F.������id,A.�걾��λ,DECODE(E.�Ա�,'��',1,'Ů',2,0),E.��������) AS ����ο�," & _
                               "D.�������," & _
                               "D.��λ," & _
                               "D.���㹫ʽ, " & _
                               "a.������ĿID, " & _
                               "h.�������,h.���鱸ע " & _
                        "FROM ����ҽ����¼ A," & _
                             "������Ŀ D," & _
                             "������Ϣ E, " & _
                             "���鱨����Ŀ h, " & _
                             "(SELECT ������id,�������� FROM ���˲��������� WHERE NVL(��,0)=0 AND ����id=[1]) F, " & _
                             "(SELECT Distinct x.ҽ��id,y.���鱸ע FROM ������Ŀ�ֲ� x,����걾��¼ y WHERE x.�걾id=y.ID And x.ҽ��id=[2]) h, " & _
                             "����������Ŀ G " & _
                        "Where A.���ID = [2] And A.���id=h.ҽ��id(+) " & _
                              "AND E.����ID=A.����id " & _
                              "AND F.������id=D.������ĿID(+) " & _
                              "AND G.ID=D.������ĿID " & _
                              "AND a.������ĿID = h.������ĿID " & _
                              "AND D.������ĿID = h.������ĿID " & _
                              " order by a.������ĿID,h.������� "
        Else
            mstrSQL = "SELECT DISTINCT C.������ĿID AS ID," & _
                               "G.������ AS ������Ŀ," & _
                               "Decode(d.�������,3,Decode(F.��������,Null,'-',''),F.��������) AS ������," & _
                               "zlGetReference(C.������ĿID,A.�걾��λ,DECODE(E.�Ա�,'��',1,'Ů',2,0),E.��������) AS ����ο�," & _
                               "D.�������," & _
                               "B.���㵥λ AS ��λ," & _
                               "D.���㹫ʽ,C.�������,h.���鱸ע " & _
                        "FROM ����ҽ����¼ A," & _
                             "������ĿĿ¼ B," & _
                             "���鱨����Ŀ C," & _
                             "������Ŀ D," & _
                             "������Ϣ E, " & _
                             "(SELECT ������id,�������� FROM ���˲��������� WHERE NVL(��,0)=0 AND ����id=[1]) F, " & _
                             "(SELECT Distinct x.ҽ��id,y.���鱸ע FROM ������Ŀ�ֲ� x,����걾��¼ y WHERE x.�걾id=y.ID And x.ҽ��id=[2]) h, " & _
                             "����������Ŀ G " & _
                        "Where A.���ID = [2] And A.���id=h.ҽ��id(+) " & _
                              "AND E.����ID=A.����id " & _
                              "AND A.������ĿID=B.ID " & _
                              "AND C.������ĿID=B.ID " & _
                              "AND F.������id(+)=C.������ĿID " & _
                              "AND D.������ĿID=C.������ĿID " & _
                              "AND G.ID=C.������ĿID Order By C.�������"
        End If
                    
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "���˲���������", "H���˲���������")
            mstrSQL = Replace(mstrSQL, "����ҽ����¼", "H����ҽ����¼")
            mstrSQL = Replace(mstrSQL, "����걾��¼", "H����걾��¼")
            mstrSQL = Replace(mstrSQL, "������Ŀ�ֲ�", "H������Ŀ�ֲ�")
        End If
        Set rs = OpenSQLRecord(mstrSQL, "���鱨��", mlng����id, mlngҽ��id)
        If rs.BOF = False Then
            
            txt.Text = zlCommFun.NVL(rs("���鱸ע").Value)
            
            Do While Not rs.EOF
                
                If zlCommFun.NVL(rs("�������").Value) = 1 Then
                    strDec = "0.0000"
                Else
                    strDec = ""
                End If
                
                If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                    vsf.Rows = vsf.Rows + 1
                End If
                
                vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID").Value)
                vsf.TextMatrix(vsf.Rows - 1, mCol.������Ŀ) = zlCommFun.NVL(rs("������Ŀ").Value)
                
                vsf.TextMatrix(vsf.Rows - 1, mCol.����ο�) = zlCommFun.NVL(rs("����ο�").Value)
                
                strTmp = zlCommFun.NVL(rs("������").Value)
                If strTmp <> "" Then
                    vsf.TextMatrix(vsf.Rows - 1, mCol.������) = IIf(strDec = "", Split(strTmp, "'")(0), Format(Split(strTmp, "'")(0), strDec))
                    If UBound(Split(strTmp, "'")) > 0 Then vsf.TextMatrix(vsf.Rows - 1, mCol.�����־) = Split(strTmp, "'")(1)
                    If UBound(Split(strTmp, "'")) > 1 Then vsf.TextMatrix(vsf.Rows - 1, mCol.����ο�) = Split(strTmp, "'")(2)
                End If
                    
                vsf.TextMatrix(vsf.Rows - 1, mCol.�������) = zlCommFun.NVL(rs("�������").Value)
                vsf.TextMatrix(vsf.Rows - 1, mCol.���㹫ʽ) = zlCommFun.NVL(rs("���㹫ʽ").Value)
                vsf.TextMatrix(vsf.Rows - 1, mCol.��λ) = zlCommFun.NVL(rs("��λ").Value)
                
                Call ApplyResultColor(vsf, vsf.Rows - 1, mCol.������, _
                    Decode(vsf.TextMatrix(vsf.Rows - 1, mCol.�����־), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
                
                rs.MoveNext
            Loop
        End If
        
        For mlngLoop = 1 To vsf.Rows - 1
            Call ApplyResultColor(vsf, mlngLoop, mCol.������, _
                Decode(vsf.TextMatrix(mlngLoop, mCol.�����־), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
        Next
        
        '�Զ����ø߶�
        If (vsf.Rows * (vsf.Body.RowHeight(0) + 15) + 30) > UserControl.Height Then
            UserControl.Height = vsf.Rows * (vsf.Body.RowHeight(0) + 15) + 30
        End If
    Else
        vsf2.Rows = 3
        vsf2.RowData(2) = 0
        vsf2.Cell(flexcpText, 2, 0, 2, vsf2.Cols - 1) = ""
        
        mstrSQL = _
            "SELECT A.ID," & _
                  "E.������ AS ϸ��," & _
                  "A.������ AS ����," & _
                  "B.������," & _
                  "B.���," & _
                  "B.����," & _
                  "B.��ɫֵ," & _
                  "A.��������,h.���鱸ע " & _
            "FROM ������ͨ��� A," & _
                 "������Ŀ C," & _
                 "����걾��¼ D," & _
                 "����ϸ�� E," & _
                 "(SELECT Distinct x.ҽ��id,y.���鱸ע FROM ������Ŀ�ֲ� x,����걾��¼ y WHERE x.�걾id=y.ID And x.ҽ��id=[1]) h, " & _
                 "����ҽ����¼ F," & _
                 "(SELECT A.ϸ�����ID," & _
                         "B.������ AS ������," & _
                         "A.���," & _
                         "DECODE(A.�������,'R','255','I','16711680','S','0','0') AS ��ɫֵ," & _
                         "DECODE(A.�������,'R','��ҩ','I','�н�','S','����','') AS ���� " & _
                  "FROM ����ҩ����� A," & _
                       "�����ÿ����� B " & _
                  "Where A.������ID = B.ID " & _
                 ") B "
                 
        mstrSQL = mstrSQL & _
            "Where A.������Ŀid = C.������ĿID(+) " & _
                "AND C.��Ŀ���(+)=2 " & _
                "AND A.��¼���� =D.������ " & _
                "AND D.ID=A.����걾ID " & _
                "AND A.ϸ��id =E.ID(+) " & _
                "AND A.ID=B.ϸ�����ID(+) " & _
                "AND (D.ҽ��id=F.ID Or D.ҽ��id=F.���ID) " & _
                "AND F.ID=[1] And f.ID=h.ҽ��id(+) ORDER BY A.ID "
                        
        mstrSQL = "SELECT ID,ϸ��,����,������,���,����,��ɫֵ,��������,���鱸ע FROM (" & mstrSQL & ") A "
                        
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "������ͨ���", "H������ͨ���")
            mstrSQL = Replace(mstrSQL, "����걾��¼", "H����걾��¼")
            mstrSQL = Replace(mstrSQL, "����ҽ����¼", "H����ҽ����¼")
            mstrSQL = Replace(mstrSQL, "����ҩ�����", "H����ҩ�����")
            mstrSQL = Replace(mstrSQL, "������Ŀ�ֲ�", "H������Ŀ�ֲ�")
        End If
        Set rs = OpenSQLRecord(mstrSQL, "���鱨��", mlngҽ��id)
        If rs.BOF = False Then
            
            txt.Text = zlCommFun.NVL(rs("���鱸ע").Value)
            
            Do While Not rs.EOF
                If Val(vsf2.RowData(vsf2.Rows - 1)) > 0 Then
                    vsf2.Rows = vsf2.Rows + 1
                End If
                
                vsf2.RowData(vsf2.Rows - 1) = zlCommFun.NVL(rs("ID").Value)
                
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.ϸ��) = zlCommFun.NVL(rs("ϸ��").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.����) = zlCommFun.NVL(rs("����").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.������) = zlCommFun.NVL(rs("������").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.���) = zlCommFun.NVL(rs("���").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.����) = zlCommFun.NVL(rs("����").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.��������) = zlCommFun.NVL(rs("��������").Value)
                
                vsf2.Cell(flexcpForeColor, vsf2.Rows - 1, mCol.���) = zlCommFun.NVL(rs("��ɫֵ").Value)
                
                rs.MoveNext
            Loop
        End If
        
        Dim lngSvrKey As Long
        Dim strSpace As String
        Dim lngLoop As Long
        
        For lngLoop = 2 To vsf2.Rows - 1
            If lngSvrKey <> Val(vsf2.RowData(lngLoop)) Then
                lngSvrKey = Val(vsf2.RowData(lngLoop))
                strSpace = IIf(strSpace = " ", "", " ")
            End If
            
            vsf2.TextMatrix(lngLoop, mCol.ϸ��) = vsf2.TextMatrix(lngLoop, mCol.ϸ��) & strSpace
            vsf2.TextMatrix(lngLoop, mCol.����) = vsf2.TextMatrix(lngLoop, mCol.����) & strSpace
            vsf2.TextMatrix(lngLoop, mCol.��������) = vsf2.TextMatrix(lngLoop, mCol.��������) & strSpace
            
        Next
        
        '�Զ����ø߶�
        If (vsf2.Rows * (vsf2.RowHeight(0) + 15) + 30) > UserControl.Height Then
            UserControl.Height = vsf2.Rows * (vsf2.RowHeight(0) + 15) + 30
        End If
        Call AppendRows(vsf2, lnX, lnY)
    End If
    
    mblnLoaded = False
    
    Exit Function
    
ErrHandle:
    
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Function
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Function

Public Sub ClearData()
    '------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    vsf.Cell(flexcpText, 1, mCol.������, vsf.Rows - 1, mCol.������) = ""
    vsf.Cell(flexcpText, 1, mCol.�����־, vsf.Rows - 1, mCol.�����־) = ""
End Sub

Public Function SaveData(lng����ID As Long, lng��ҳID As Long, lng����ID As Long, strReturnSQL As String, strError As String) As Boolean
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSQL() As String
    Dim strValue As String
    
    ReDim Preserve strSQL(0 To vsf.Rows)
    
    If mblnCommon = False Then
        strReturnSQL = ""
        SaveData = True
        Exit Function
    End If
    
    Call vsf_AfterEdit(vsf.Row, vsf.Col)
    
'    strSQL(0) = "ZL_���˲�������_DELETE(" & lng����ID & ")"
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            Select Case Val(Left(vsf.TextMatrix(lngLoop, mCol.�����־), 1))
                Case 3
                    vsf.TextMatrix(lngLoop, mCol.�����־) = "��"
                Case 2
                    vsf.TextMatrix(lngLoop, mCol.�����־) = "��"
                Case 4
                    vsf.TextMatrix(lngLoop, mCol.�����־) = "�쳣"
                Case 5
                    vsf.TextMatrix(lngLoop, mCol.�����־) = "����"
                Case 6
                    vsf.TextMatrix(lngLoop, mCol.�����־) = "����"
    '            Case Else
    '                vsf.TextMatrix(Row, Col) = ""
            End Select
            
            strValue = vsf.TextMatrix(lngLoop, mCol.������) & "''" & vsf.TextMatrix(lngLoop, mCol.�����־) & "''" & vsf.TextMatrix(lngLoop, mCol.����ο�)
            
            strSQL(lngLoop) = "ZL_���˲���������_SAVE(" & lng����ID & "," & _
                                                        lngLoop & "," & _
                                                        "2,'" & _
                                                        vsf.TextMatrix(lngLoop, mCol.������Ŀ) & "'," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        Val(vsf.RowData(lngLoop)) & "," & _
                                                        vsf.TextMatrix(lngLoop, mCol.�������) & ",'" & _
                                                        vsf.TextMatrix(lngLoop, mCol.��λ) & "','" & _
                                                        strValue & "'" & _
                                                        ")"
        End If
    Next
        
    strTmp = ""
    For lngLoop = 0 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then
            If strTmp = "" Then
                strTmp = strSQL(lngLoop)
            Else
                strTmp = strTmp & Chr(9) & strSQL(lngLoop)
            End If
        End If
    Next
    
    '����SQL���
    strReturnSQL = strTmp
    
    SaveData = True
    
End Function

Private Sub UserControl_Initialize()

    '��ʼ���ؼ�����
    
    On Error GoTo ErrHandle
    
                
    Exit Sub
    
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InDesign() As Boolean
    
    '���ܣ��жϵ�ǰ���г����Ƿ���VB�Ĺ��̻�����
    
    On Error Resume Next
    
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
    
End Function

Private Sub UserControl_InitProperties()
    '��ʼ���˲���Ϊ0
    mlng����id = 0
    mDispMode = True
    mblnMoved = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDispMode = PropBag.ReadProperty("DispMode", True)
    mblnMoved = PropBag.ReadProperty("DataMoved", False)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", BorderStyleSettings.flexBorderNone)
End Sub

Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Let Locked(ByVal vData As Boolean)
    MsgBox "a"
End Property

Private Sub UserControl_Resize()
    On Error Resume Next
    
    With shp
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height - picComment.Height
    End With
    
    With vsf
        .Left = 15
        .Top = 15
        .Width = UserControl.Width - .Left - 15
        .Height = UserControl.Height - .Top - 15
    End With
    
    With vsf2
        .Left = vsf.Left
        .Top = vsf.Top
        .Width = vsf.Width
        .Height = vsf.Height
    End With
    
    picComment.Move 0, shp.Top + shp.Height, shp.Width
    txt.Move txt.Left, -15, picComment.Width - txt.Left, picComment.Height + 15
    
    Call AppendRows(vsf2, lnX, lnY)
End Sub

Private Sub UserControl_Terminate()
    If rsTmp.State = adStateOpen Then rsTmp.Close
    Set rsTmp = Nothing
    
    On Error Resume Next
    
    Set mobjParentObject = Nothing
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DispMode", mDispMode, True)
    Call PropBag.WriteProperty("DataMoved", mblnMoved, False)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, BorderStyleSettings.flexBorderNone)
End Sub

Private Sub UserControl_Show()
    Dim objCtl As Control
         
    'ֻ������ʱ��ʾ
    
    On Error Resume Next
    
    If Ambient.UserMode = True And InDesign = False Then
        If mDispMode Then
            For Each objCtl In Controls
                If UCase(TypeName(objCtl)) <> UCase("ImageList") Then
                    objCtl.Enabled = False
                End If
            Next
        End If
    End If
    
    If mblnLoaded = False Then InitData
        
    
End Sub

Public Property Get Text() As String
    'Ϊÿһ���ؼ������ı�ת������
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSvrKey As Long
    
    'ͨ���û���������ݵõ�ת���ı�
    strTmp = "���鱨�棺" & vbCrLf
    If mblnCommon Then
        For lngLoop = 0 To vsf.Rows - 1
            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.������Ŀ) & Space(50), 1, 50)
            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.������) & Space(20), 1, 20)
            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.�����־) & Space(20), 1, 20)
            strTmp = strTmp & vsf.TextMatrix(lngLoop, mCol.����ο�) & vbCrLf
            If lngLoop = 0 Then strTmp = strTmp & "------------------------------------------------------------------------------------------" & vbCrLf
            If lngLoop = vsf.Rows - 1 Then strTmp = strTmp & "------------------------------------------------------------------------------------------" & vbCrLf
        Next
    Else
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.ϸ��) & Space(40), 1, 40)
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.����) & Space(20), 1, 20)
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.������) & Space(40), 1, 40)
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.���) & Space(20), 1, 20)
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.����) & Space(20), 1, 20)
        strTmp = strTmp & vsf2.TextMatrix(1, mCol.��������) & vbCrLf
        strTmp = strTmp & "-----------------------------------------------------------------------------------------------------" & vbCrLf
        
        For lngLoop = 2 To vsf2.Rows - 1
            
            If strSvrKey <> vsf2.RowData(lngLoop) Then
                
                strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.ϸ��) & Space(40), 1, 40)
                strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.����) & Space(20), 1, 20)
                
            End If
            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.������) & Space(40), 1, 40)
            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.���) & Space(20), 1, 20)
            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.����) & Space(20), 1, 20)
            
            If strSvrKey <> vsf2.RowData(lngLoop) Then
                
                strSvrKey = vsf2.RowData(lngLoop)
                strTmp = strTmp & vsf2.TextMatrix(lngLoop, mCol.��������)
                
            End If
            strTmp = strTmp & vbCrLf
        Next
        
        strTmp = strTmp & "-----------------------------------------------------------------------------------------------------" & vbCrLf
    End If
    
    Text = strTmp
    
End Property

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    
    UserControl.Parent.CallBack_GotFocus
    
End Sub


Private Sub vsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
'    mblnChangeEdit = True
'    Call AdjustEnableState
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strReference As String
    Dim LngCount As Long
    
    If Col = mCol.�����־ Then
        Select Case Val(Left(vsf.TextMatrix(Row, mCol.�����־), 1))
            Case 3
                vsf.TextMatrix(Row, Col) = "��"
            Case 2
                vsf.TextMatrix(Row, Col) = "��"
            Case 4
                vsf.TextMatrix(Row, Col) = "�쳣"
            Case 5
                vsf.TextMatrix(Row, Col) = "����"
            Case 6
                vsf.TextMatrix(Row, Col) = "����"
'            Case Else
'                vsf.TextMatrix(Row, Col) = ""
        End Select
        Call ApplyResultColor(vsf, Row, mCol.������, Decode(vsf.TextMatrix(Row, mCol.�����־), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
    End If
    
    If Col = mCol.������ And Val(vsf.TextMatrix(Row, mCol.�������)) <> 2 Then
        
        '����ȱʡ�Ľ����־
        vsf.TextMatrix(Row, mCol.�����־) = Format(CalcDefaultFlag(Trim(vsf.TextMatrix(Row, Col)), Trim(vsf.TextMatrix(Row, mCol.����ο�)), Val(vsf.TextMatrix(Row, mCol.�������))), "0.0000")
        
        '���ݽ��Ӧ����ɫ��־
        Call ApplyResultColor(vsf, Row, mCol.������, Decode(vsf.TextMatrix(Row, mCol.�����־), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
                                
                                
        '�Զ����������Ŀ���
        For mlngLoop = 1 To vsf.Rows - 1
            If Trim(vsf.TextMatrix(mlngLoop, mCol.���㹫ʽ)) <> "" Then
                
                vsf.TextMatrix(mlngLoop, Col) = Format(CalcExpress(vsf, Trim(vsf.TextMatrix(mlngLoop, mCol.���㹫ʽ))), "0.0000")
                
                '����ȱʡ�Ľ����־
                vsf.TextMatrix(mlngLoop, mCol.�����־) = CalcDefaultFlag(Trim(vsf.TextMatrix(mlngLoop, Col)), Trim(vsf.TextMatrix(mlngLoop, mCol.����ο�)), Val(vsf.TextMatrix(mlngLoop, mCol.�������)))
        
                '���ݽ��Ӧ����ɫ��־
                Call ApplyResultColor(vsf, Row, mCol.������, Decode(vsf.TextMatrix(Row, mCol.�����־), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
            End If
        Next
        
    End If

'    mblnChangeEdit = True
'    Call AdjustEnableState
End Sub



Private Sub vsf_BeforeComboList(ByVal OldCol As Long, ByVal NewCol As Long, ComboList As String, Cancel As Boolean)
    On Error Resume Next
    
    If mblnModify = False Then Exit Sub
    
    '1-������2-ƫ�͡�3-ƫ�ߡ�4-����
    '1:����,2:���֣�3��������(+-)
    If NewCol = mCol.�����־ Then
        Select Case Val(vsf.TextMatrix(vsf.Row, mCol.�������))
            Case 1  '����
                ComboList = "1-����|2-ƫ��|3-ƫ��"
'                ComboList = "|��|��"
            Case 2  '����
                ComboList = "1-����|4-�쳣"
'                ComboList = "1-����|4-�쳣"
            Case 3  '�붨��
                ComboList = "1-����|2-ƫ��|3-ƫ��|4-�쳣"
        End Select
    ElseIf NewCol = mCol.������ Then
        ComboList = "" '"|-|+|--|++|+-"
    End If
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Col = mCol.������
    Cancel = True
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    '����������͵�
    On Error Resume Next
    
    If mblnModify = False Then Exit Sub
    
    If NewCol = mCol.������ Then
        If Trim(vsf.TextMatrix(NewRow, mCol.���㹫ʽ)) <> "" Then
            vsf.EditMode(NewCol) = 0
        Else
            vsf.EditMode(NewCol) = 1
        End If
    
    
        Select Case Val(vsf.TextMatrix(NewRow, mCol.�������))
        Case 2
            vsf.ComboList(mCol.������) = "..."
            vsf.VsfComboList = "..."
        Case 3
            vsf.ComboList(mCol.������) = " "
            vsf.VsfComboList = " |-|+|--|++|+-"
        Case Else
            vsf.ComboList(mCol.������) = ""
            vsf.VsfComboList = ""
        End Select
        
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case ShowOpenList(, False)
    Case 0
        'û��ƥ�����Ŀ
        'MsgBox "û���ҵ���ƥ��Ľ����", vbInformation, gstrSysName
        
    Case 1
        'ѡȡ��һ����Ŀ
'        mblnChangeEdit = True
'        Call AdjustEnableState
    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    
    Dim strSvrText As String
    
    If mblnModify = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        '����2-�����͵����
        If Val(vsf.TextMatrix(Row, mCol.�������)) <> 2 Then Exit Sub
        
        If InStr(vsf.EditText, "'") > 0 Then
            Cancel = True
            Exit Sub
        End If

        strSvrText = vsf.EditText
        Select Case ShowOpenList(vsf.EditText, True)
        Case 0
            'û��ƥ�����Ŀ
            vsf.Cell(flexcpData, Row, Col) = strSvrText
        Case 1
            'ѡȡ��һ����Ŀ
'            mblnChangeEdit = True
'            Call AdjustEnableState
        Case 2
            'ȡ���˱���ѡ��
            Cancel = True

            vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
            vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)

        End Select
    Else
'        mblnChangeEdit = True
'        Call AdjustEnableState
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If mblnModify = False Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    
    Select Case Val(vsf.TextMatrix(vsf.Row, mCol.�������))
    Case 1
        KeyAscii = FilterKeyAscii(KeyAscii, 2)
    Case Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End Select
    
End Sub

Private Sub vsf2_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf2, lnX, lnY)
End Sub

Private Sub vsf2_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf2, lnX, lnY)
End Sub
'�����Ƿ�ת��
Public Property Get DataMoved() As Boolean
    DataMoved = mblnMoved
End Property

Public Property Let DataMoved(ByVal vNewValue As Boolean)
    mblnMoved = vNewValue
End Property


