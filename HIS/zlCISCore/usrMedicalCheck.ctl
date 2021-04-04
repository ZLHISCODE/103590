VERSION 5.00
Begin VB.UserControl usrMedicalCheck 
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   ScaleHeight     =   4710
   ScaleWidth      =   7350
   Begin zl9CISCore.VsfGrid vsf 
      Height          =   1530
      Left            =   1740
      TabIndex        =   0
      Top             =   345
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   2699
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000003&
      Height          =   435
      Left            =   270
      Top             =   495
      Width           =   585
   End
End
Attribute VB_Name = "usrMedicalCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private rsTmp As New ADODB.Recordset
Private mlng����id As Long                      '��紫��
Private mlngҽ��id As Long                      '��紫��
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

    ��Ŀ = 1
    ���
    ��λ
    ����
    ��ֵ��
    ������
    ��ʼֵ
    ��ʾ��
    ��ֵ����
End Enum

Private Enum COLOR
    
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    
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

'Private Function ShowOpenList(Optional strText As String, Optional blnWhere As Boolean = False) As Byte
'    '-----------------------------------------------------------------------------------------
'    '����:���б�ṹ�����Ƽ���걾����
'    '����:������2;�ɹ�����1;ȡ������0
'    '-----------------------------------------------------------------------------------------
'    Dim strLvw As String
'    Dim sglX As Single
'    Dim sglY As Single
'    Dim rs As New ADODB.Recordset
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    strLvw = "����,900,0,1;ȡֵ,1800,0,0;�����־,900,0,0"
'
'    ShowOpenList = 2
'
'    strSQL = "SELECT ROWNUM AS ID,����,ȡֵ,DECODE(�����־,1,'1-����',2,'2-ƫ��',3,'3-ƫ��',4,'4-����',5,'5-����','') AS �����־ FROM ������Ŀȡֵ A WHERE ��Ŀid=[1]"
'    If blnWhere Then
'        strSQL = strSQL & " AND (A.���� Like [2] OR A.ȡֵ Like [2])"
'    End If
'
'    Set rs = OpenSQLRecord(strSQL, "������", CLng(Val(vsf.RowData(vsf.Row))), "%" & strText & "%")
'
'    If rs.BOF Then
'
'        ShowOpenList = 0
'
'        Exit Function
'    End If
'
'    If rs.RecordCount = 1 And blnWhere Then GoTo Over
'
'    Call CalcPosition(sglX, sglY, vsf)
'
'    If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 3600, 4500, "����ȡֵѡ��", "����±���ѡ��һ��ȡֵ") Then
'        GoTo Over
'    End If
'
'    Exit Function
'
'Over:
''    vsf.EditText = zlCommFun.Nvl(rs("ȡֵ").Value)
''    vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("ȡֵ").Value)
''    vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("ȡֵ").Value)
''    vsf.TextMatrix(vsf.Row, mCol.�����־) = zlCommFun.Nvl(rs("�����־").Value)
'
'    ShowOpenList = 1
'
'    Exit Function
'
'errHand:
'    If ErrCenter = 1 Then Resume
'End Function

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
        
    strSQL = "SELECT DISTINCT A.ִ��״̬,D.��Ŀ��� FROM ����ҽ������ A,����ҽ����¼ B,���鱨����Ŀ C,������Ŀ D WHERE B.������Ŀid=C.������Ŀid AND C.������Ŀid=D.������Ŀid AND A.ҽ��id=B.ID AND B.���id=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rs = OpenSQLRecord(strSQL, "���鱨��ר��ֽ", mlngҽ��id)
    If rs.BOF = False Then mblnModify = (rs("ִ��״̬").Value <> 1)
                    
    vsf.Visible = False
        
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "��Ŀ", 1500, 1
        .NewColumn "���", 3300, 1, , 1
        .NewColumn "��λ", 600, 1
        .NewColumn "����", 0, 1
        .NewColumn "��ֵ��", 0, 1
        .NewColumn "������", 0, 1
        .NewColumn "��ʼֵ", 0, 1
        .NewColumn "��ʾ��", 0, 1
        .NewColumn "��ֵ����", 0, 1
    
        .NewColumn "", 255, 4
        
        .ExtendLastCol = True
        .Body.Appearance = flexFlat
        
        .Body.BorderStyle = flexBorderNone
        .Body.BackColorFixed = .Body.BackColor
        .Body.ColHidden(mCol.����) = True
        .Body.ColHidden(mCol.��ֵ��) = True
        .Body.ColHidden(mCol.������) = True
        .Body.ColHidden(mCol.��ʼֵ) = True
        .Body.ColHidden(mCol.��ʾ��) = True
        .Body.ColHidden(mCol.��ֵ����) = True
        
        .FixedCols = 1
        
        .Cell(flexcpFontBold, 0, 0, 0, vsf.Cols - 1) = True
        
        .SelectMode = True
        
        .Visible = True
    End With
    
    If mblnModify = False Then
        vsf.EditMode(mCol.��Ŀ) = 0
        vsf.EditMode(mCol.���) = 0
        vsf.EditMode(mCol.��λ) = 0
        vsf.ComboList(mCol.���) = ""

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
                                        
    
    '�������
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    
    '��ȡ��װ������
    
    mstrSQL = "Select 1 From ���˲�����¼ a,���˲������� b Where a.��д���� Is Not Null And a.id=b.������¼id and b.ID=[1]"
    Set rs = OpenSQLRecord(mstrSQL, "����鱨��", mlng����id)
    If rs.BOF Then mlng����id = 0
    
    If mlng����id > 0 Then

        mstrSQL = "Select A.ID,a.������ As ��Ŀ,b.�������� As ���,a.��λ,a.����,a.��ֵ��,a.������,a.��ʼֵ,a.��ʾ��,a.��ֵ���� " & _
                    "From ����������Ŀ a,���˲��������� b " & _
                    "Where a.ID=b.������id and b.����id=[1] "
                                
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "���˲���������", "H���˲���������")
            mstrSQL = Replace(mstrSQL, "����ҽ����¼", "H����ҽ����¼")
        End If
        Set rs = OpenSQLRecord(mstrSQL, "����鱨��", mlng����id)
    Else
        
        mstrSQL = "Select a.ID,a.������ As ��Ŀ,Decode(a.��ʼֵ,Null,a.��ֵ����,a.��ʼֵ) As ���,a.��λ,a.����,a.��ֵ��,a.������,a.��ʼֵ,a.��ʾ��,a.��ֵ���� " & _
                    "From ����������Ŀ a,���������� b,����Ԫ��Ŀ¼ c,����ҽ����¼ d " & _
                    "Where c.����=-1 and c.id=b.Ԫ��id and b.��=d.������Ŀid and a.id=b.������id and d.ID=[1] " & _
                    "Order By b.�ؼ��� "
        
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "����ҽ����¼", "H����ҽ����¼")
        End If
        Set rs = OpenSQLRecord(mstrSQL, "����鱨��", mlngҽ��id)
        
    End If
    
    If rs.BOF = False Then
        Do While Not rs.EOF
                    
            If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.��Ŀ) = zlCommFun.NVL(rs("��Ŀ").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.��λ) = zlCommFun.NVL(rs("��λ").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.���) = zlCommFun.NVL(rs("���").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.����) = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.��ֵ��) = zlCommFun.NVL(rs("��ֵ��").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.������) = zlCommFun.NVL(rs("������").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.��ʼֵ) = zlCommFun.NVL(rs("��ʼֵ").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.��ʾ��) = zlCommFun.NVL(rs("��ʾ��").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.��ֵ����) = zlCommFun.NVL(rs("��ֵ����").Value)
                                            
            rs.MoveNext
        Loop
    End If
    
    '�Զ����ø߶�
    If (vsf.Rows * (vsf.Body.RowHeight(0) + 15) + 30) < UserControl.Height Then
        UserControl.Height = vsf.Rows * (vsf.Body.RowHeight(0) + 15) + 30
    End If
    
    vsf.Cell(flexcpForeColor, 1, mCol.���, vsf.Rows - 1, mCol.���) = COLOR.��ɫ
    
    Call vsf.Body.AutoSize(mCol.���, mCol.���)
    
    mblnLoaded = False
    
    Exit Function
    
ErrHandle:
    
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Function
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
    
End Function

Public Sub ClearData()
    '------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    vsf.Cell(flexcpText, 1, mCol.���, vsf.Rows - 1, mCol.���) = ""
    
End Sub

Public Function SaveData(lng����ID As Long, lng��ҳID As Long, lng����ID As Long, strReturnSQL As String, strError As String) As Boolean
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSQL() As String
    Dim strValue As String
    
    ReDim Preserve strSQL(0 To vsf.Rows)
    
'    strSQL(0) = "ZL_���˲�������_DELETE(" & lng����ID & ")"
    
    If mblnModify Then
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                
                strValue = vsf.TextMatrix(lngLoop, mCol.���)
                
                strSQL(lngLoop) = "ZL_���˲���������_SAVE(" & lng����ID & "," & _
                                                            lngLoop & "," & _
                                                            "2,'" & _
                                                            vsf.TextMatrix(lngLoop, mCol.��Ŀ) & "'," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            Val(vsf.RowData(lngLoop)) & "," & _
                                                            "0,'" & _
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
    End If
    
    SaveData = True
    
End Function

Private Sub UserControl_Initialize()

    '��ʼ���ؼ�����
    
    On Error GoTo ErrHandle
    
    vsf.ComboEdit = True
    vsf.SelEdit = True
    vsf.Body.WordWrap = True
    
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
        .Height = UserControl.Height
    End With
    
    With vsf
        .Left = 15
        .Top = 15
        .Width = UserControl.Width - .Left - 15
        .Height = UserControl.Height - .Top - 15
    End With

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
'
'    'ͨ���û���������ݵõ�ת���ı�
'    strTmp = "���鱨�棺" & vbCrLf
'    If mblnCommon Then
'        For lngLoop = 0 To vsf.Rows - 1
'            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.������Ŀ) & Space(50), 1, 50)
'            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.������) & Space(20), 1, 20)
'            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.�����־) & Space(20), 1, 20)
'            strTmp = strTmp & vsf.TextMatrix(lngLoop, mCol.����ο�) & vbCrLf
'            If lngLoop = 0 Then strTmp = strTmp & "------------------------------------------------------------------------------------------" & vbCrLf
'            If lngLoop = vsf.Rows - 1 Then strTmp = strTmp & "------------------------------------------------------------------------------------------" & vbCrLf
'        Next
'    Else
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.ϸ��) & Space(40), 1, 40)
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.����) & Space(20), 1, 20)
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.������) & Space(40), 1, 40)
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.���) & Space(20), 1, 20)
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.����) & Space(20), 1, 20)
'        strTmp = strTmp & vsf2.TextMatrix(1, mCol.��������) & vbCrLf
'        strTmp = strTmp & "-----------------------------------------------------------------------------------------------------" & vbCrLf
'
'        For lngLoop = 2 To vsf2.Rows - 1
'
'            If strSvrKey <> vsf2.RowData(lngLoop) Then
'
'                strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.ϸ��) & Space(40), 1, 40)
'                strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.����) & Space(20), 1, 20)
'
'            End If
'            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.������) & Space(40), 1, 40)
'            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.���) & Space(20), 1, 20)
'            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.����) & Space(20), 1, 20)
'
'            If strSvrKey <> vsf2.RowData(lngLoop) Then
'
'                strSvrKey = vsf2.RowData(lngLoop)
'                strTmp = strTmp & vsf2.TextMatrix(lngLoop, mCol.��������)
'
'            End If
'            strTmp = strTmp & vbCrLf
'        Next
'
'        strTmp = strTmp & "-----------------------------------------------------------------------------------------------------" & vbCrLf
'    End If
'
    Text = strTmp
    
End Property

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    
    UserControl.Parent.CallBack_GotFocus
    
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Trim(vsf.TextMatrix(Row, mCol.���)) = "" Then
        vsf.TextMatrix(Row, mCol.���) = vsf.TextMatrix(Row, mCol.��ֵ����)
    End If
    
    Call vsf.Body.AutoSize(mCol.���, mCol.���)
    
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Col = mCol.���
    Cancel = True
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    Dim strTmp As String
    Dim aryTmp As Variant
    
    On Error Resume Next
    
    If mblnModify = False Then Exit Sub

    If NewCol = mCol.��� Then
        
        '0-�ı�,1-����,2-����,3-��ѡ,4-��ѡ;5-ָ��(����Ŀ������ֵ��ĳ�����ݱ����ͼ���������������ݲ��ṩ)
        If Val(vsf.TextMatrix(NewRow, mCol.����)) = 1 Then
            Select Case Val(vsf.TextMatrix(NewRow, mCol.��ʾ��))
'            Case 2, 4
'                strTmp = vsf.TextMatrix(NewRow, mCol.��ֵ��)
'
'                aryTmp = Split(strTmp, ";")
'                strTmp = " |" & Join(aryTmp, "|")
'
'                vsf.ComboList(mCol.���) = strTmp
            Case 2, 4, 3    '��ѡ
                vsf.ComboList(mCol.���) = "..."
            Case Else
                vsf.ComboList(mCol.���) = ""
            End Select
        Else
            vsf.ComboList(mCol.���) = ""
        End If
        
        vsf.VsfComboList = vsf.ComboList(mCol.���)
    End If

End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim sglX As Single
    Dim sglY As Single
    Dim strText As String
    Dim strTmp As String
    Dim aryTmp As Variant
    Dim lngLoop As Long
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    rs.Fields.Append "ID", adBigInt
    rs.Fields.Append "ĩ��", adBigInt
    rs.Fields.Append "ѡ��", adBigInt
    rs.Fields.Append "����", adVarChar, 100
    rs.Open
    
    strTmp = vsf.TextMatrix(Row, mCol.��ֵ��)
    aryTmp = Split(strTmp, ";")
    
    strText = ";" & vsf.TextMatrix(Row, mCol.���) & ";"
    
    For lngLoop = LBound(aryTmp) To UBound(aryTmp)
        
        If InStr(strText, ";" & aryTmp(lngLoop) & ";") > 0 Then
            rs.AddNew
            rs("ID").Value = lngLoop
            rs("ĩ��").Value = 1
            rs("ѡ��").Value = 1
            rs("����").Value = CStr(aryTmp(lngLoop))
                        
        Else
            
            rs.AddNew
            rs("ID").Value = lngLoop
            rs("ĩ��").Value = 1
            rs("ѡ��").Value = 0
            rs("����").Value = CStr(aryTmp(lngLoop))
            
        End If
        
    Next
        
    If rs.RecordCount = 0 Then Exit Sub
    
    rs.MoveFirst
    
    Call CalcPosition(sglX, sglY, vsf)
    
    Dim blnMuli As Boolean
    
    If Val(vsf.TextMatrix(Row, mCol.��ʾ��)) = 3 Then
        '��ѡ
        blnMuli = True
    End If
    
    If frmSelectDialog.ShowSelect(Nothing, 2, rs, "����,3300,0,0", "�������ѡ������Ŀ,Ȼ��س���˫���˳�", sglX + 60, sglY + 30, 6000, vsf.Body.ColWidth(Col), 300, , "�����ѡ��", , False, blnMuli) Then
        
        vsf.TextMatrix(Row, Col) = ""
        
        If blnMuli Then
            rs.Filter = ""
            rs.Filter = "ѡ��=1"
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    vsf.TextMatrix(vsf.Row, vsf.Col) = vsf.TextMatrix(vsf.Row, vsf.Col) & zlCommFun.NVL(rs("����").Value) & ";"
                    rs.MoveNext
                Loop
                
                If vsf.TextMatrix(Row, Col) <> "" Then vsf.TextMatrix(Row, Col) = Mid(vsf.TextMatrix(Row, Col), 1, Len(vsf.TextMatrix(Row, Col)) - 1)
                            
            End If
        Else
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
        End If
        
        Call vsf.Body.AutoSize(mCol.���, mCol.���)
    End If
End Sub

Private Sub vsf_ChangeEdit(ByVal Row As Long, ByVal Col As Long)
    vsf.TextMatrix(Row, Col) = vsf.EditText
    Call vsf.Body.AutoSize(mCol.���, mCol.���)
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If mblnModify = False Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    
    '0-��ֵ��1-���֣�2-���ڣ�(3-�߼�)
    
    Select Case Val(vsf.TextMatrix(vsf.Row, mCol.����))
    Case 0
        KeyAscii = FilterKeyAscii(KeyAscii, 2)
    Case Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End Select
    
    
End Sub

'�����Ƿ�ת��
Public Property Get DataMoved() As Boolean
    DataMoved = mblnMoved
End Property

Public Property Let DataMoved(ByVal vNewValue As Boolean)
    mblnMoved = vNewValue
End Property

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Dim aryTmp As Variant
    
    If Col <> mCol.��� Then Exit Sub
    If Trim(vsf.EditText) = "" Then Exit Sub
    
    '0-��ֵ��1-���֣�2-���ڣ�(3-�߼�)
    If Val(vsf.TextMatrix(Row, mCol.����)) = 0 Then
        If vsf.TextMatrix(Row, mCol.��ֵ��) <> "" Then
            
            aryTmp = Split(Trim(vsf.TextMatrix(Row, mCol.��ֵ��)), ";")
            If Val(vsf.EditText) < aryTmp(0) Then
                vsf.EditText = ""
                Cancel = True
            End If
            
            If UBound(aryTmp) > 0 Then
                If Val(vsf.EditText) > aryTmp(1) Then
                    vsf.EditText = ""
                    Cancel = True
                End If
            End If
        End If
    End If
End Sub
