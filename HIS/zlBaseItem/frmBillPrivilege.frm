VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillPrivilege 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ݲ���Ȩ������"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmBillPrivilege.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3990
      TabIndex        =   11
      Top             =   240
      Width           =   1100
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "����Ȩ��"
      Height          =   2385
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   3675
      Begin VB.TextBox txt������� 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         Height          =   270
         Left            =   2280
         TabIndex        =   10
         Top             =   1980
         Width           =   1035
      End
      Begin VB.ComboBox cmb��Ա 
         Height          =   300
         Left            =   1410
         TabIndex        =   2
         Text            =   "cmb��Ա"
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   810
         Width           =   1935
      End
      Begin VB.TextBox txtUD 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   2670
         TabIndex        =   6
         Text            =   "0"
         Top             =   1260
         Width           =   390
      End
      Begin VB.CheckBox chk�޸� 
         Caption         =   "׼��������˵���(&T)"
         Height          =   210
         Left            =   270
         TabIndex        =   8
         Top             =   1665
         Value           =   1  'Checked
         Width           =   2040
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   300
         Left            =   3060
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1260
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtUD(1)"
         BuddyDispid     =   196614
         BuddyIndex      =   1
         OrigLeft        =   3105
         OrigTop         =   1260
         OrigRight       =   3345
         OrigBottom      =   1560
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������Ľ������(&M)"
         Height          =   180
         Index           =   3
         Left            =   285
         TabIndex        =   9
         Top             =   2025
         Width           =   1890
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������(&S)"
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   3
         Top             =   870
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����Ա(&N)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   420
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����������ݵ���ʷ����(&D)"
         Height          =   180
         Index           =   2
         Left            =   270
         TabIndex        =   5
         Top             =   1320
         Width           =   2250
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   12
      Top             =   750
      Width           =   1100
   End
End
Attribute VB_Name = "frmBillPrivilege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr���� As String, mstr��ԱID As String, mstr���� As String
Private mlng���� As Long, mlng���� As Long, mbln�޸����� As Boolean
Private mdbl������� As Double

Dim mblnChange As Boolean     '�Ƿ�ı���
Dim mblnOk As Boolean
Dim mstrLike As String
Private Sub chk�޸�_Click()
    mblnChange = True
End Sub

Private Sub cmb����_Click()
    mblnChange = True
    If Mid(cmb����.Text, 1, 1) = 2 Or Mid(cmb����.Text, 1, 1) = 4 Or Mid(cmb����.Text, 1, 1) = 5 Or Mid(cmb����.Text, 1, 1) = 9 Then
        Me.txt�������.Enabled = True
    Else
        Me.txt�������.Text = "0.00"
        Me.txt�������.Enabled = False
    End If
End Sub

Private Sub cmb��Ա_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim intIdx As Integer
    
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    
    If cmb��Ա.ItemData(cmb��Ա.ListIndex) = -1 And Visible Then
        strSQL = "Select ID,����,���� From ��Ա�� Where ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null Order by ����"

        vRect = GetControlRect(cmb��Ա.hwnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "����Ա", , , , , , True, vRect.Left, vRect.Top, cmb��Ա.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cmb��Ա, rsTmp!ID)
            If intIdx <> -1 Then
                cmb��Ա.ListIndex = intIdx
            Else
                cmb��Ա.AddItem Nvl(rsTmp!����) & "-" & rsTmp!����, cmb��Ա.ListCount - 1
                cmb��Ա.ItemData(cmb��Ա.NewIndex) = rsTmp!ID
                cmb��Ա.ListIndex = cmb��Ա.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�в���Ա���ݣ����ȵ�����/��Ա���������á�", vbInformation, gstrSysName
            End If
            '�ָ������е���Ա(������Click)
            intIdx = SeekCboIndex(cmb��Ա, cmb��Ա.Tag)
            Call zlControl.CboSetIndex(cmb��Ա.hwnd, intIdx)
        End If
    Else
        cmb��Ա.Tag = cmb��Ա.Text
    End If
End Sub

Private Function SeekCboIndex(objCbo As Object, varData As Variant) As Long
'���ܣ���ItemData��Text����ComboBox������ֵ
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.Type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '�Ⱦ�ȷ����
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '��ģ������
            For i = 0 To objCbo.ListCount - 1
                If InStr(objCbo.List(i), varData) > 0 And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    Else
        If varData <> 0 Then
            For i = 0 To objCbo.ListCount - 1
                If objCbo.ItemData(i) = varData Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    End If
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function
Private Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function
Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Private Sub cmb��Ա_GotFocus()
    If cmb��Ա.Style = 0 Then
        Call zlControl.TxtSelAll(cmb��Ա)
    End If
End Sub

Private Sub cmb��Ա_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cmb��Ա.Style = 2 And cmb��Ա.ListIndex <> -1 Then
            cmb��Ա.ListIndex = -1
        End If
    End If
End Sub


Private Sub cmb��Ա_KeyPress(KeyAscii As Integer)
'    Dim lngIdx As Long
'
'    lngIdx = MatchIndex(cmb��Ա.hwnd, KeyAscii)
'    If lngIdx <> -2 Then cmb��Ա.ListIndex = lngIdx
    
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cmb��Ա.Locked And cmb��Ա.Style = 2 Then
            lngIdx = zlControl.CboMatchIndex(cmb��Ա.hwnd, KeyAscii)
            If lngIdx = -1 And cmb��Ա.ListCount > 0 Then lngIdx = 0
            cmb��Ա.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cmb��Ա_Validate(Cancel As Boolean)
    '���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cmb��Ա.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cmb��Ա.Text = "" Then cmb��Ա.Tag = "": Exit Sub '������
    
    strInput = UCase(NeedName(cmb��Ա.Text))
    strSQL = "Select ID,����,���� From ��Ա�� Where (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) Order by ����"
    strSQL = Replace(UCase(strSQL), UCase("Order by"), " And (Upper(���) Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2]) Order by")
    
    On Error GoTo errH
    vRect = GetControlRect(cmb��Ա.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����Ա", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cmb��Ա.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = SeekCboIndex(cmb��Ա, rsTmp!ID)
        If intIdx <> -1 Then
            cmb��Ա.ListIndex = intIdx
        Else
            cmb��Ա.AddItem Nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����, cmb��Ա.ListCount - 1
            cmb��Ա.ItemData(cmb��Ա.NewIndex) = rsTmp!ID
            cmb��Ա.ListIndex = cmb��Ա.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�Ĳ���Ա��", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
End Sub


Private Sub txtUD_Change(Index As Integer)
    If Index = 1 Then
        If Val(txtUD(Index).Text) > 100 Then txtUD(Index).Text = 100
        If Val(txtUD(Index).Text) < 0 Then txtUD(Index).Text = 0
    End If
End Sub

Private Sub txtUD_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtUD(1))
End Sub


Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        KeyAscii = 0
    End If
End Sub


Private Sub txt�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(Me.txt�������.Text) = "" Then
            Me.txt�������.Text = "0.00"
        End If
        If Not IsNumeric(Me.txt�������.Text) Then
            MsgBox "����Ľ���ʽ����ȷ��"
            Me.txt�������.SetFocus
            Exit Sub
        ElseIf Val(Me.txt�������.Text) > 10000000 Then
            MsgBox "���ܳ���7λ������"
            Me.txt�������.SetFocus
            Exit Sub
        End If
        Me.txt�������.Text = Format(Me.txt�������.Text, "0.00")
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If

End Sub


Private Sub ud_Change()
    mblnChange = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    
    mstr���� = Replace(GetTextFromCombo(cmb��Ա, True), "'", "")
    mstr��ԱID = cmb��Ա.ItemData(cmb��Ա.ListIndex)
    
    mstr���� = Mid(cmb����.Text, 3)
    mlng���� = Left(cmb����.Text, 1)
    mlng���� = Val(txtUD(1).Text)
    mbln�޸����� = (chk�޸�.Value = 1)
    mdbl������� = Val(txt�������.Text)
    
    mblnOk = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    If cmb��Ա.ListIndex < 0 Then
        MsgBox "��ѡ�����Ա��", vbInformation, gstrSysName
        cmb��Ա.SetFocus
        Exit Function
    End If
    If Trim(cmb��Ա.Text) = "" Then
        MsgBox "��ѡ�����Ա��", vbInformation, gstrSysName
        cmb��Ա.SetFocus
        Exit Function
    End If
    If cmb����.Text = "" Then
        MsgBox "��ѡ������ĵ������͡�", vbInformation, gstrSysName
        cmb����.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(Me.txt�������.Text) Then
        MsgBox "����Ľ���ʽ����ȷ��"
        Me.txt�������.SetFocus
        Exit Function
    ElseIf Val(Me.txt�������.Text) > 10000000 Then
        MsgBox "���ܳ���7λ������"
        Me.txt�������.SetFocus
        Exit Function
    End If
    Me.txt�������.Text = Format(Me.txt�������.Text, "0.00")
    
'    If ud.Value = 0 And chk�޸�.Value = 1 Then
'        MsgBox "�Ե��ݵĲ������ںͲ����˶�û�����ƣ����豣�档", vbInformation, gstrSysName
'        chk�޸�.SetFocus
'        Exit Function
'    End If
    
    IsValid = True
End Function

Public Function �༭Ȩ��(str���� As String, str��ԱID As String, str���� As String, lng���� As Long, lng���� As Long, bln�޸����� As Boolean, dbl������� As Double _
                        , frmParent As Form) As Boolean
'���ܣ���Ϊ�ӿں���
    Dim rsTemp As New ADODB.Recordset, str���� As String
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID,����,���� From ��Ա�� Where ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    cmb��Ա.Clear
    Do Until rsTemp.EOF
        If IsNull(rsTemp("����")) Then
            str���� = zlStr.GetCodeByVB(rsTemp("����"))
        Else
            str���� = rsTemp("����")
        End If
        cmb��Ա.AddItem str���� & "-" & rsTemp("����")
        cmb��Ա.ItemData(cmb��Ա.NewIndex) = rsTemp("ID")
        rsTemp.MoveNext
    Loop
    
    cmb����.Clear
    
    If glngSys \ 100 = 8 Then
        'ҩ��ϵͳ����ĵ���������
        cmb����.AddItem "2.�շѵ�"
        cmb����.AddItem "8.��Ա��"
    Else
        cmb����.AddItem "1.�Һŵ���"
        cmb����.AddItem "2.�շѵ�"
        cmb����.AddItem "3.���۵�"
        cmb����.AddItem "4.�������"
        cmb����.AddItem "5.סԺ����"
        cmb����.AddItem "6.Ԥ����"
        cmb����.AddItem "7.���ʵ���"
        cmb����.AddItem "8.���￨"
        cmb����.AddItem "9.����"
    End If
    
    mstr��ԱID = str��ԱID
    SetComboByText cmb��Ա, str����, True
    '�޸ı��2779
    If cmb��Ա.List(cmb��Ա.ListCount - 1) = "" Then
        'ɾ���Ǹ�����
        cmb��Ա.RemoveItem cmb��Ա.ListCount - 1
    End If
    '----------------------------------
    If cmb��Ա.ListIndex >= 0 Then
        '���������ԭ�б��в����ڵ���Ա����������ID
        If cmb��Ա.ItemData(cmb��Ա.ListIndex) = 0 Then cmb��Ա.ItemData(cmb��Ա.ListIndex) = Val(str��ԱID)
    End If
    
    ud.Value = lng����
    chk�޸�.Value = IIF(bln�޸�����, 1, 0)
    txt�������.Text = Format(dbl�������, "0.00")
    SetComboByText cmb����, lng����, False, "."
    
    mblnChange = False
    mblnOk = False
    frmBillPrivilege.Show vbModal, frmParent
    
    
    If mblnOk = True Then
        str���� = mstr����
        str��ԱID = mstr��ԱID
        str���� = mstr����
        lng���� = mlng����
        lng���� = mlng����
        bln�޸����� = mbln�޸�����
        dbl������� = mdbl�������
    End If
    �༭Ȩ�� = mblnOk
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
