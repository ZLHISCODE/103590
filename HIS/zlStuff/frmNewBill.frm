VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ϵ����쳣����"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7170
   Icon            =   "frmNewBill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7170
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtShow 
      Height          =   2535
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmNewBill.frx":014A
      Top             =   1200
      Width           =   6735
   End
   Begin VB.CommandButton Cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   8
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   7065
   End
   Begin MSComctlLib.ProgressBar prg������ 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4290
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4800
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtPatiId 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   795
      Width           =   2415
   End
   Begin VB.ComboBox cboDept 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   240
      Picture         =   "frmNewBill.frx":0155
      Top             =   150
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "¼�벡����Ϣ�������ڲ�������ѯ�Ƿ����δ�����ķ��ϵ��ݣ�������ھ��Զ����²�����"
      Height          =   420
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   6060
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lblPatiId 
      AutoSize        =   -1  'True
      Caption         =   "����š�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   840
   End
   Begin VB.Menu mnuPati 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuPatiItem 
         Caption         =   "�����(&0)"
         Index           =   0
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "סԺ��(&1)"
         Index           =   1
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "���￨��(&2)"
         Index           =   2
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "����(&3)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmNewBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjServiceCall As Object           '����
Private mintType As Integer '1-������id���ң�2-����������

Private Enum FindType
    ����� = 0
    סԺ��
    ���￨��
    ����
End Enum

Private mCol��ʾ��Ϣ As New Collection
Private mstrShow As String
Public Sub ShowForm(frmMain As Form, Optional ByVal intType As Integer = 1)
    '�������
    mintType = intType
    
    Me.Show vbModal, frmMain
End Sub

Private Sub cmdOK_Click()
    Dim intResult As Integer
    Dim colInput As New Collection, colPati As New Collection
    Dim i As Integer
    Dim rsPati As adodb.Recordset, rsSelPati As adodb.Recordset
    Dim strErrMsg As String
    Dim strPatiOut As String, strPatiIDs As String
    Dim varList As Variant  '����Ԫ��
    Dim cllErrMsg As Collection '������Ϣ������Ա(Array(��������,��������,������Ϣ,����S(N)))
        
    On Error GoTo ErrHandle
    If mintType = 1 Then
        If txtPatiId.Text = "" Then
            MsgBox "�������벡��" & Replace(lblPatiId.Caption, "��", "") & "��", vbInformation, gstrSysName
            zlControl.ControlSetFocus txtPatiId
            Exit Sub
        End If
        
        If Val(txtPatiId.Tag) = 0 Then
            colInput.Add Null, "pati_id"
            colInput.Add Null, "outpatient_num"
            colInput.Add Null, "inpatient_num"
            colInput.Add Null, "pati_wardarea_id"
            colInput.Add Null, "pati_bed"
            colInput.Add Null, "pati_deptid"
            colInput.Add Null, "pati_name"
            colInput.Add Null, "pati_vcard_no"
    
            '��������
            Select Case Val(lblPatiId.Tag)
            Case FindType.�����
                If Not IsNumeric(txtPatiId.Text) Then
                    MsgBox "�������Ч�����������룡", vbInformation, gstrSysName
                    zlControl.ControlSetFocus txtPatiId: zlControl.TxtSelAll txtPatiId
                    Exit Sub
                End If
                
                colInput.Remove ("outpatient_num")
                colInput.Add Val(txtPatiId.Text), "outpatient_num"
            
            Case FindType.סԺ��
                If Not IsNumeric(txtPatiId.Text) Then
                    MsgBox "סԺ����Ч�����������룡", vbInformation, gstrSysName
                    zlControl.ControlSetFocus txtPatiId: zlControl.TxtSelAll txtPatiId
                    Exit Sub
                End If
                
                'ͨ��סԺ���Ҳ���ID
                If zlSplitService_GetPatiId(mobjServiceCall, 1342, txtPatiId.Text, strPatiOut) = False Then Exit Sub
                If Val(strPatiOut) = 0 Then Exit Sub
                
                '����id
                colInput.Remove ("pati_id")
                colInput.Add Val(strPatiOut), "pati_id"
            Case FindType.���￨��
                colInput.Remove ("pati_vcard_no")
                colInput.Add txtPatiId.Text, "pati_vcard_no"
            Case FindType.����
                colInput.Remove ("pati_name")
                colInput.Add txtPatiId.Text, "pati_name"
            End Select
            
            If zlSplitService_GetPatiName(mobjServiceCall, 1342, colInput, colPati) = False Then Exit Sub
            If colPati.Count = 0 Then
                MsgBox "δ�ҵ���Ӧ�Ĳ�����Ϣ��", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If colPati.Count = 1 Then
                txtPatiId.Text = IIf(colPati(1)("_pati_dept_name") = "", "", colPati(1)("_pati_dept_name") & "-") & colPati(1)("_pati_name")
                txtPatiId.Tag = Val(colPati(1)("_pati_id"))
            Else
                '���ض�����¼ʱ
                Set rsPati = New adodb.Recordset
                With rsPati
                    If .State = 1 Then .Close
                    .Fields.Append "����id", adDouble, 18, adFldIsNullable
                    .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
                    .Fields.Append "סԺ��", adDouble, 18, adFldIsNullable
                    .Fields.Append "����", adLongVarChar, 30, adFldIsNullable
                    .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
                    .Fields.Append "����id", adDouble, 18, adFldIsNullable
                    .Fields.Append "����", adLongVarChar, 30, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .Open
                    
                    For i = 1 To colPati.Count
                        .AddNew
                        
                        !����ID = colPati(i)("_pati_id")
                        !�������� = colPati(i)("_pati_name")
                        !סԺ�� = colPati(i)("_inpatient_num")
                        !���� = colPati(i)("_pati_wardarea_name")
                        !���� = colPati(i)("_pati_bed")
                        !����id = colPati(i)("_pati_dept_id")
                        !���� = colPati(i)("_pati_dept_name")
                        
                        .Update
                    Next
                End With
                
                If zlDatabase.zlShowListSelect(Me, 100, 1342, txtPatiId, rsPati, True, "", "����ID,����ID", rsSelPati) = False Then Exit Sub
            
                rsSelPati.Filter = ""
                If rsSelPati.RecordCount = 0 Then Exit Sub
                
                txtPatiId.Text = IIf(rsSelPati!���� = "", "", rsSelPati!���� & "-") & rsSelPati!��������
                txtPatiId.Tag = rsSelPati!����ID
            End If
        End If

        '��鲢���²�������
        If Val(txtPatiId.Tag) = 0 Then Exit Sub
        
        intResult = ExecuteDataSync(Val(txtPatiId.Tag), cllErrMsg)
        strErrMsg = GetErrMsg(cllErrMsg)
        Select Case intResult
        Case 0
            MsgBox "���ˡ�" & Mid(txtPatiId.Text, InStr(txtPatiId.Text, "-") + 1) & "��δ�����ķ��ϵ��������²�����ɣ�", vbInformation, gstrSysName
        Case 1
            MsgBox "���ˡ�" & Mid(txtPatiId.Text, InStr(txtPatiId.Text, "-") + 1) & "��������δ�����ķ��ϵ��ݣ�", vbInformation, gstrSysName
        Case 2
            MsgBox "�ڼ�鲡�ˡ�" & Mid(txtPatiId.Text, InStr(txtPatiId.Text, "-") + 1) & "���Ƿ����δ�����ķ��ϵ���ʱ���ִ���" & _
                    IIf(strErrMsg = "", "", vbCrLf & vbCrLf & strErrMsg), vbInformation, gstrSysName
        Case 3
            MsgBox "���ˡ�" & Mid(txtPatiId.Text, InStr(txtPatiId.Text, "-") + 1) & "��δ�����Ĳ��ַ��ϵ������²���ʱʧ�ܣ�" & _
                IIf(strErrMsg = "", "", vbCrLf & vbCrLf & strErrMsg), vbInformation, gstrSysName
        End Select
        zlControl.ControlSetFocus txtPatiId: zlControl.TxtSelAll txtPatiId
        
        Call Show����������Ϣ
        
        Exit Sub
    End If
    
    '2-����������
    If Val(cboDept.ItemData(cboDept.ListIndex)) <= 0 Then
        MsgBox "����ѡ��һ��������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�������ȡ��������id
    If zlSplitService_GetPatiByRange(mobjServiceCall, 1342, Val(cboDept.ItemData(cboDept.ListIndex)), colPati) = False Then Exit Sub
    If colPati.Count = 0 Then
        MsgBox "��" & cboDept.Text & "��������δ�����ķ��ϵ��ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ݲ���id��鲢ͬ���쳣����
    For Each varList In colPati
        strPatiIDs = strPatiIDs & "," & Val(varList("_pati_id"))
    Next
    
    intResult = ExecuteDataSync(Mid(strPatiIDs, 2), cllErrMsg)
    strErrMsg = GetErrMsg(cllErrMsg)
    Select Case intResult
    Case 0
        MsgBox "��" & cboDept.Text & "��δ�����ķ��ϵ��������²�����ɣ�", vbInformation, gstrSysName
    Case 1
        MsgBox "��" & cboDept.Text & "��������δ�����ķ��ϵ��ݣ�", vbInformation, gstrSysName
    Case 2
        MsgBox "�ڼ�顾" & cboDept.Text & "���Ƿ����δ�����ķ��ϵ���ʱ���ִ���" & _
            IIf(strErrMsg = "", "", vbCrLf & vbCrLf & strErrMsg), vbInformation, gstrSysName
    Case 3
        MsgBox "��" & cboDept.Text & "��δ�����Ĳ��ַ��ϵ������²���ʱʧ�ܣ�" & _
            IIf(strErrMsg = "", "", vbCrLf & vbCrLf & strErrMsg), vbInformation, gstrSysName
    End Select
    
    Call Show����������Ϣ
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetErrMsg(ByVal cllErrMsg As Collection, Optional ByVal bln��ʾ���� As Boolean) As String
    '��ȡ������Ϣ
    '��Σ�
    '   cllErrMsg-������Ϣ������Ա(Array(��������,��������,������Ϣ,����S(N)))  �������ͣ�2-�ٴ���ͬ�����ʧ�ܣ�1-������ͬ�����ʧ�ܣ�0-��������
    Dim i As Long, strMsg As String, strErrInfo As String
    Dim lngCount As Long, bytErrType As Byte, strInfo As String
    
    If cllErrMsg Is Nothing Then Exit Function
    
    strMsg = "": lngCount = 0
    For i = 1 To cllErrMsg.Count
        bytErrType = cllErrMsg(i)(0)
        
        strErrInfo = cllErrMsg(i)(2)
        If InStr(UCase(strErrInfo), "[ZLSOFT]") > 0 Then strErrInfo = Split(strErrInfo, "[ZLSOFT]")(1)
        
        strInfo = ""
        If strErrInfo <> "" Then
            If lngCount > 5 Then '����5��ʡ�Ժű�ʾ
                strMsg = strMsg & vbCrLf & "����"
                Exit For
            End If
            
            strInfo = (lngCount + 1) & "��"
            If cllErrMsg(i)(1) <> "" And bln��ʾ���� Then strInfo = strInfo & cllErrMsg(i)(1) & " "
            If bytErrType = 2 Then
                strInfo = strInfo & "[" & cllErrMsg(i)(3) & "] �޷�ͬ���������ҽ�����·��͡�ԭ��"
            ElseIf bytErrType = 1 Then
                strInfo = strInfo & "[" & cllErrMsg(i)(3) & "] �޷�ͬ���������Ϸ������¼Ƿѡ�ԭ��"
            Else
                strInfo = strInfo & "ͬ��ʧ�ܣ������ԡ�ԭ��"
            End If
            strInfo = strInfo & strErrInfo
            
            If strInfo <> "" And InStr(vbCrLf & strMsg & vbCrLf, vbCrLf & strInfo & vbCrLf) = 0 Then
                strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & strInfo
                lngCount = lngCount + 1
            End If
        End If
    Next
    If lngCount = 1 Then strMsg = Mid(strMsg, 3)
    
    GetErrMsg = strMsg
End Function

Private Sub Cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If mintType = 2 Then
        lblDept.Visible = True
        cboDept.Visible = True
        lblPatiId.Visible = False
        txtPatiId.Visible = False
        Call LoadDept
    End If

    'ʵ��������
    Call zlSercieCall_Ini(mobjServiceCall)
    mobjServiceCall.InitService gcnOracle, gstrDBUser, glngSys, glngModul
End Sub

Private Sub LoadDept()
    Dim rsTemp As adodb.Recordset, strSQL As String
    
    On Error GoTo ErrHandle
    cboDept.Clear
    cboDept.Tag = ""
    
    strSQL = _
        " Select b.���� As վ������, b.��� As վ��,A.����||'-'||A.���� ����,A.ID" & _
        " From ���ű� A, Zlnodelist B " & _
        " Where a.վ�� = b.���(+) And A.ID in (Select ����ID From ��������˵�� Where �������� in('����','�ٴ�') And ������� IN(2,3))" & _
        "           And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
        " Order By a.վ��, a.���� || '-' || a.���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����")
    Do While Not rsTemp.EOF
        cboDept.AddItem rsTemp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTemp!Id
        rsTemp.MoveNext
    Loop
    If cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call zlSercieCall_Unload(mobjServiceCall)
End Sub


Private Sub lblPatiId_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        PopupMenu mnuPati, 2, lblPatiId.Left + lblPatiId.Width - 30, lblPatiId.Top
    End If
End Sub

Private Sub mnuPatiItem_Click(index As Integer)
    Dim i As Integer
    
    lblPatiId.Tag = index
    txtPatiId.Text = ""
    txtPatiId.MaxLength = 0
    
    Select Case index
        Case FindType.�����
            lblPatiId.Caption = "����š�"
            lblPatiId.Tag = FindType.�����
        Case FindType.סԺ��
            lblPatiId.Caption = "סԺ�š�"
            lblPatiId.Tag = FindType.סԺ��
        Case FindType.���￨��
            lblPatiId.Caption = "���￨�š�"
            lblPatiId.Tag = FindType.���￨��
        Case FindType.����
            lblPatiId.Caption = "������"
            lblPatiId.Tag = FindType.����
    End Select
    
    For i = 0 To mnuPatiItem.Count - 1
        mnuPatiItem(i).Checked = (i = index)
    Next
End Sub

Private Sub RefrashProgress(Optional ByVal lngValue As Long, Optional ByVal bytMode As Byte = 1, Optional ByVal lngMaxValue As Long)
    'ˢ�½�����ʾ
    '���:
    '   bytMode-���ͣ�0-ˢ����Ϣ��1-��ʼ����ʾ��2-��ֹ��ʾ
    On Error GoTo ErrHandler
    Select Case bytMode
    Case 0
        Me.MousePointer = vbHourglass
        prg������.Visible = True
        prg������.Value = 0
        prg������.Max = lngMaxValue
    Case 1
        prg������.Value = lngValue
    Case 2
        prg������.Visible = False
        Me.MousePointer = vbDefault
    End Select
    Exit Sub
ErrHandler:
    prg������.Visible = False
    Me.MousePointer = vbDefault
End Sub

Private Function ExecuteDataSync(ByVal strPatiIDs As String, ByRef cllErrMsg_Out As Collection) As Integer
    'ִ���쳣����ͬ��
    '��Σ�
    '   strPatiIDs-����ID�������Ӣ�Ķ��ŷָ�
    '���Σ�
    '   cllErrMsg_Out-������Ϣ������Ա(Array(��������,��������,������Ϣ,����S(N)))
    '���أ�0-����δ�����ķ��ϵ��ݣ�������ȫ��������1-������δ�����ķ��ϵ��ݣ�2-��������3-����δ�����ķ��ϵ��ݣ��������²����ɹ�
    '˵��:
    '   1.�ٴ���ͬ���쳣����������+����"����ͬ��
    '   2.�������쳣���������ݡ�����ͬ��
    Dim cllCisErrData As Collection, cllExseErrData As Collection, cllPatiData As Collection
    Dim cllOrderSendItem As Collection, cllPatiBillItem As Collection
    Dim i As Long, lngCount As Long, lngSccussCount As Long, strErrMsg As String
    Dim cllPati As Collection, bytErrType As Byte, strNos As String
    
    On Error GoTo ErrHandler
    Set cllErrMsg_Out = New Collection
    
    Me.MousePointer = vbHourglass
    '1.���ݲ���IDȡҽ������
    ExecuteDataSync = GetCisSyncErrData(strPatiIDs, cllCisErrData, strErrMsg)
    If ExecuteDataSync = 2 Then
        cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg)
        Me.MousePointer = vbDefault: Exit Function
    End If
    
    Me.MousePointer = vbHourglass
    '2.ȡ��������
    ExecuteDataSync = GetExseSyncErrData(strPatiIDs, cllCisErrData, cllExseErrData, strErrMsg)
    If ExecuteDataSync = 2 Or ExecuteDataSync = 1 Then
        cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg)
        Me.MousePointer = vbDefault: Exit Function
    End If

    Me.MousePointer = vbHourglass
    '3.��ȡ������Ϣ����ݣ��������ڣ����֤��
    If GetPatiData(cllExseErrData, cllPatiData, strErrMsg) = False Then
        ExecuteDataSync = 2
        cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg)
        Me.MousePointer = vbDefault: Exit Function
    End If
    
    Call RefrashProgress(, 0, cllCisErrData.Count)
    lngCount = 0: lngSccussCount = 0
    
    '4.�����ٴ���ͬ���쳣��ͬ����� cllExseErrData �Ƴ�
    For Each cllOrderSendItem In cllCisErrData
        If ExecuteCisErrDataSync(cllOrderSendItem, cllExseErrData, cllPatiData, strErrMsg, bytErrType, strNos) = False Then
            If ExistsColObject(cllPatiData, "_" & cllOrderSendItem("����ID")) Then
                Set cllPati = cllPatiData("_" & cllOrderSendItem("����ID"))
                cllErrMsg_Out.Add Array(bytErrType, cllPati("����"), strErrMsg, strNos)
            Else
                cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg, strNos)
            End If
        Else
            lngSccussCount = lngSccussCount + 1
        End If
        bytErrType = 0
        
        lngCount = lngCount + 1
        Call RefrashProgress(lngCount)
    Next
    If cllCisErrData.Count <> lngSccussCount Then ExecuteDataSync = 3
    
    Call RefrashProgress(, 0, cllExseErrData.Count)
    lngCount = 0: lngSccussCount = 0
    
    '5.����������ͬ���쳣
    For Each cllPatiBillItem In cllExseErrData
        If ExecuteExseErrDataSync(cllPatiBillItem, cllPatiData, strErrMsg, bytErrType, strNos) = False Then
            If (cllPatiBillItem("��������")) = 3 Then '���ʱ�
                cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg, strNos)
            Else
                If ExistsColObject(cllPatiData, "_" & cllPatiBillItem("����ID")) Then
                    Set cllPati = cllPatiData("_" & cllPatiBillItem("����ID"))
                    cllErrMsg_Out.Add Array(bytErrType, cllPati("����"), strErrMsg, strNos)
                Else
                    cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg, strNos)
                End If
            End If
        Else
            lngSccussCount = lngSccussCount + 1
        End If
        bytErrType = 0
        
        lngCount = lngCount + 1
        Call RefrashProgress(lngCount)
    Next
    If cllExseErrData.Count <> lngSccussCount Then ExecuteDataSync = 3
    
    Call RefrashProgress(, 2)
    Me.MousePointer = vbDefault
    Exit Function
ErrHandler:
    cllErrMsg_Out.Add Array(bytErrType, "", err.Description)
    Me.MousePointer = vbDefault
    Call RefrashProgress(, 2)
    ExecuteDataSync = 2
End Function

Private Function ExecuteCisErrDataSync(ByVal cllOrderSendItem As Collection, ByRef cllExseErrData As Collection, _
    ByVal cllPatiData As Collection, ByRef strErrMsg As String, ByRef bytErrType As Byte, ByRef strNos As String) As Boolean
    'ִ���ٴ����쳣����ͬ��
    '��Σ�
    '   cllOrderSendItem-����ҽ�����ͼ�¼����Ա(����ID,��ҳID,�Һ�ID,�Һŵ���,���ͺ�,OrderList)
    '           |-cllOrderList-ҽ����Ϣ�б�=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-ҽ����Ϣ����Ա(ҽ��ID,ҽ����Ч,������־,�Ƽ�����)=cllOrderList(_ҽ��ID)
    '           |-cllExseBillList-���õ����б�=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-���õ�����Ϣ����Ա(������Դ,��������,���ݺ�)=cllExseBillList(_������Դ_��������_���ݺ�)
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ
    '   cllExseErrData=������ͬ���쳣���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiBillItem-���˵��ݼ�¼����Ա(��������,������Դ,[����ID,��ҳID,����,�Ա���,�Ա�,����,���˿���ID,���˲���ID],BillLists)�����У��������е�Ԫ�ؼ��ʱ�ʱ��
    '           |-cllBillLists-������Ϣ��=cllPatiBillItem(BillLists)
    '               |-cllBillItem-������Ϣ����Ա(������Դ,NO,�շѱ�־,������,��������ID,������������,
    '                                                          ����ҽʦID,����ҽʦ,����Ա����,����Ա���,�Ǽ�ʱ��,DetailList)=cllBillLists(_������Դ_��������_���ݺ�)
    '                   |-cllDetailList-������ϸ��=cllBillItem(DetailList)
    '                       |-cllDetailItem-ÿ����ϸ���ݼ�����Ա([����ID,��ҳID,����,�Ա���,�Ա�,����,���˲���ID,���˿���ID],
    '                                 ����ID,���,�ⷿID,�Ƿ񱸻�����,����,����ID,Ӥ�����,ҽ��ID,����,����,�ۼ�,���۽��,ժҪ)�����У��������е�Ԫ�ؼ��ʱ�ʱ����
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ,4-��죻������Դ��1-����,2-סԺ��
    '                 �������ͣ�0�Ϳ�-��ͨ,1-����,2-����,3-����,4-��һ,5-�����շѱ�־��0-δ�շѻ���ʻ���,1-���շѻ����
    '   cllPatiData=������Ϣ���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiItem-������Ϣ����Ա(����ID,��������,���֤��,���")=cllPatiData(_����ID)
    '���Σ�
    '   strErrMsg=������Ϣ
    '   bytErrType=�������ͣ�2-�ٴ���ͬ�����ʧ�ܣ�0-��������
    '   strNos=�漰�ĵ��ݺţ���ʽ��A001,A002,...
    '����:ִ�гɹ�����True��ִ��ʧ�ܷ���False
    Dim strJson As String, strListJson As String, strOrders As String
    Dim cllOrderList As Collection, cllOrderItem As Collection
    Dim cllExseBillList As Collection, cllExseBillItem As Collection
    Dim strNewBillCheckJson As String, strNewBillJson As String, strSyncJson As String
    Dim blnTrans As Boolean, strKey As String
    Dim cllPatiBillItem As Collection, cllBillLists As Collection
    
    On Error GoTo ErrHandler
    strErrMsg = "": bytErrType = 0: strNos = ""
    If cllOrderSendItem Is Nothing Then ExecuteCisErrDataSync = True: Exit Function
    
    Set cllOrderList = cllOrderSendItem("OrderList")
    Set cllExseBillList = cllOrderSendItem("ExseBillList")
    
    If GetNewBillJson_Cis(cllOrderSendItem, cllExseErrData, cllPatiData, _
        strNewBillCheckJson, strNewBillJson, strErrMsg, strNos) = False Then GoTo MoveExseNOsHandler
    
    If strNewBillJson = "" Then '�޷��õ��ݣ�����
        ExecuteCisErrDataSync = True
        GoTo MoveExseNOsHandler
    End If
    
    bytErrType = 2
    If mobjServiceCall.CallService("Zl_�������۳���_Check", strNewBillCheckJson, , , , False, , , , True) = False Then
        strErrMsg = "���ò����µĴ������ʧ�ܣ�": GoTo MoveExseNOsHandler
    End If
    bytErrType = 0
    
    '��ȡ�ٴ���ͬ������JSON
    'Zl_CisSvr_UpdateSyncState
    '  --���ܣ�ͬ�����¼����
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --      order_list[]
    '  --          order_id          N 1 ҽ��id
    '  --          send_no           N 1 ���ͺ�
    '  --          sign_type         N 1 ���ñ��¼�����ͣ�˵����1-���������¼,2-��� ����ҩƷͬ�����,3-��� ��������ͬ�����
    strListJson = ""
    For Each cllOrderItem In cllOrderList
        If InStr("," & strOrders & ",", ",3:" & cllOrderItem("ҽ��ID") & ",") = 0 Then
            strOrders = strOrders & ",3:" & cllOrderItem("ҽ��ID")
            
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("order_id", cllOrderItem("ҽ��ID"), 1)
            strJson = strJson & "," & GetJsonNodeString("send_no", cllOrderSendItem("���ͺ�"), 1)
            strJson = strJson & "," & GetJsonNodeString("sign_type", 3, 1)
            strListJson = strListJson & ",{" & strJson & "}"
        End If
    Next
    strSyncJson = "{""input"":{""order_list"":[" & Mid(strListJson, 2) & "]}}"
    
    gcnOracle.BeginTrans: blnTrans = True
        If mobjServiceCall.CallService("Zl_ҩƷ�շ���¼_Newstuffbill", strNewBillJson, , , , False, , , , True) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            strErrMsg = "���ù��̲���ҩƷ��������ʧ�ܣ�": GoTo MoveExseNOsHandler
        End If
        
        If mobjServiceCall.CallService("Zl_CisSvr_UpdateSyncState", strSyncJson, , , , False, , , , True) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            strErrMsg = "���÷����޸�ҽ��ͬ����־ʧ�ܣ�": GoTo MoveExseNOsHandler
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    'ִ�гɹ������ݼ�����ʾ����
    mCol��ʾ��Ϣ.Add mCol��ʾ��Ϣ.Count + 1 & " " & mstrShow
    
    ExecuteCisErrDataSync = True
    
MoveExseNOsHandler:
    '�Ƴ�ҽ���漰�ķ��õ���
    If cllExseBillList Is Nothing Then Exit Function
    For Each cllExseBillItem In cllExseBillList
        strKey = "_" & cllExseBillItem("������Դ") & "_" & cllExseBillItem("��������") & "_" & cllExseBillItem("���ݺ�")
        Dim i As Long
        For i = cllExseErrData.Count To 1 Step -1
            Set cllPatiBillItem = cllExseErrData(i)
            Set cllBillLists = cllPatiBillItem("BillLists")
            If ExistsColObject(cllBillLists, strKey) Then
                cllBillLists.Remove strKey
                If cllBillLists.Count = 0 Then cllExseErrData.Remove i
                Exit For
            End If
        Next
    Next
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    strErrMsg = err.Description
    GoTo MoveExseNOsHandler
End Function

Private Function ExecuteExseErrDataSync(ByVal cllPatiBillItem As Collection, ByVal cllPatiData As Collection, _
    ByRef strErrMsg As String, ByRef bytErrType As Byte, ByRef strNos As String) As Boolean
    'ִ���ٴ����쳣����ͬ��
    '��Σ�
    '   cllPatiBillItem-���˵��ݼ�¼����Ա(��������,������Դ,[����ID,��ҳID,����,�Ա���,�Ա�,����,���˿���ID,���˲���ID],BillLists)�����У��������е�Ԫ�ؼ��ʱ�ʱ��
    '           |-cllBillLists-������Ϣ��=cllPatiBillItem(BillLists)
    '               |-cllBillItem-������Ϣ����Ա(������Դ,NO,�շѱ�־,������,��������ID,������������,
    '                                                          ����ҽʦID,����ҽʦ,����Ա����,����Ա���,�Ǽ�ʱ��,DetailList)=cllBillLists(_������Դ_��������_���ݺ�)
    '                   |-cllDetailList-������ϸ��=cllBillItem(DetailList)
    '                       |-cllDetailItem-ÿ����ϸ���ݼ�����Ա([����ID,��ҳID,����,�Ա���,�Ա�,����,���˲���ID,���˿���ID],
    '                                 ����ID,���,�ⷿID,�Ƿ񱸻�����,����,����ID,Ӥ�����,ҽ��ID,����,����,�ۼ�,���۽��,ժҪ)�����У��������е�Ԫ�ؼ��ʱ�ʱ����
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ,4-��죻������Դ��1-����,2-סԺ��
    '                 �������ͣ�0�Ϳ�-��ͨ,1-����,2-����,3-����,4-��һ,5-�����շѱ�־��0-δ�շѻ���ʻ���,1-���շѻ����
    '   cllPatiData=������Ϣ���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiItem-������Ϣ����Ա(����ID,��������,���֤��,���")=cllPatiData(_����ID)
    '���Σ�
    '   strErrMsg=������Ϣ
    '   bytErrType=�������ͣ�1-������ͬ�����ʧ�ܣ�0-��������
    '   strNos=�漰�ĵ��ݺţ���ʽ��A001,A002,...
    '����:ִ�гɹ�����True��ִ��ʧ�ܷ���False
    Dim strSyncJson As String, str����ids As String
    Dim strNewBillCheckJson As String, strNewBillJson As String
    Dim cllBillLists As Collection, cllBillItem As Collection
    Dim cllDetailItem  As Collection, cllDetailList As Collection
    Dim blnTrans As Boolean
    
    On Error GoTo ErrHandler
    strErrMsg = "": bytErrType = 0: strNos = ""
    If cllPatiBillItem Is Nothing Then ExecuteExseErrDataSync = True: Exit Function
    
    Set cllBillLists = cllPatiBillItem("BillLists")
    
    If GetNewBillJson_Exse(cllPatiBillItem, cllPatiData, strNewBillCheckJson, strNewBillJson, strErrMsg, strNos) = False Then Exit Function
    
    bytErrType = 1
    If mobjServiceCall.CallService("Zl_�������۳���_Check", strNewBillCheckJson, , , , False, , , , True) = False Then
        strErrMsg = "���ò����µĴ������ʧ�ܣ�": Exit Function
    End If
    bytErrType = 0
    
    str����ids = ""
    '��ȡ������ͬ������JSON
    'Zl_Exsesvr_Sync_Update
    '      ---------------------------------------------------------------------------
    '  --���ܣ�����ͬ������ռǷ�ͬ����־����NO�򰴷���ID��
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --    sign_type           N 1 ��־���ͣ�0-�Ƿ�ͬ����־,1-ת��ͬ����־
    '  --    detail_ids  C  1  ������ϸid��(����id��),֧�ֶ��id���á�,���ָ�
    '  --    bill_list[]
    '  --      billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
    '  --      rcp_no                 C   1 ����No
    For Each cllBillItem In cllBillLists
        Set cllDetailList = cllBillItem("DetailList")
        For Each cllDetailItem In cllDetailList
            str����ids = str����ids & "," & cllDetailItem("����ID")
        Next
    Next
    strSyncJson = "{""input"":{""sign_type"":0,""detail_ids"":""" & Mid(str����ids, 2) & """}}"
    
    gcnOracle.BeginTrans: blnTrans = True
        If mobjServiceCall.CallService("Zl_ҩƷ�շ���¼_Newstuffbill", strNewBillJson, , , , False, , , , True) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            strErrMsg = "���ù��̲���ҩƷ��������ʧ�ܣ�": Exit Function
        End If
        
        If mobjServiceCall.CallService("Zl_Exsesvr_Sync_Update", strSyncJson, , , , False, , , , True) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            strErrMsg = "���÷����޸ļǷ�ͬ����־ʧ�ܣ�": Exit Function
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    'ִ�гɹ������ݼ�����ʾ����
    mCol��ʾ��Ϣ.Add mCol��ʾ��Ϣ.Count + 1 & " " & mstrShow
    
    ExecuteExseErrDataSync = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    strErrMsg = err.Description
End Function

Private Sub Show����������Ϣ()
    Dim strShow As String
    Dim i As Integer
    
    If Not mCol��ʾ��Ϣ Is Nothing Then
        If mCol��ʾ��Ϣ.Count > 0 Then
            For i = 1 To mCol��ʾ��Ϣ.Count
                strShow = IIf(strShow = "", "", strShow & vbCrLf) & mCol��ʾ��Ϣ(i)
            Next
        End If
    End If
    
    If strShow = "" Then
        txtShow.Text = "���������²�����"
    Else
        txtShow.Text = "�������������²�����" & vbCrLf & strShow
    End If
End Sub
Private Function GetNewBillJson_Cis(ByVal cllOrderSendItem As Collection, ByVal cllExseErrData As Collection, ByVal cllPatiData As Collection, _
    ByRef strNewBillCheckJson_Out As String, ByRef strNewBillJson_Out As String, ByRef strErrMsg As String, ByRef strNos As String) As Boolean
    'ִ���ٴ����쳣����ͬ��
    '��Σ�
    '   cllOrderSendItem-����ҽ�����ͼ�¼����Ա(����ID,��ҳID,�Һ�ID,�Һŵ���,���ͺ�,OrderList)
    '           |-cllOrderList-ҽ����Ϣ�б�=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-ҽ����Ϣ����Ա(ҽ��ID,ҽ����Ч,������־,�Ƽ�����)=cllOrderList(_ҽ��ID)
    '           |-cllExseBillList-���õ����б�=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-���õ�����Ϣ����Ա(������Դ,��������,���ݺ�)=cllExseBillList(_������Դ_��������_���ݺ�)
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ
    '   cllExseErrData=������ͬ���쳣���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiBillItem-���˵��ݼ�¼����Ա(��������,������Դ,[����ID,��ҳID,����,�Ա���,�Ա�,����,���˿���ID,���˲���ID],BillLists)�����У��������е�Ԫ�ؼ��ʱ�ʱ��
    '           |-cllBillLists-������Ϣ��=cllPatiBillItem(BillLists)
    '               |-cllBillItem-������Ϣ����Ա(������Դ,NO,�շѱ�־,������,��������ID,������������,
    '                                                          ����ҽʦID,����ҽʦ,����Ա����,����Ա���,�Ǽ�ʱ��,DetailList)=cllBillLists(_������Դ_��������_���ݺ�)
    '                   |-cllDetailList-������ϸ��=cllBillItem(DetailList)
    '                       |-cllDetailItem-ÿ����ϸ���ݼ�����Ա([����ID,��ҳID,����,�Ա���,�Ա�,����,���˲���ID,���˿���ID],
    '                                 ����ID,���,�ⷿID,�Ƿ񱸻�����,����,����ID,Ӥ�����,ҽ��ID,����,����,�ۼ�,���۽��,ժҪ)�����У��������е�Ԫ�ؼ��ʱ�ʱ����
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ,4-��죻������Դ��1-����,2-סԺ��
    '                 �������ͣ�0�Ϳ�-��ͨ,1-����,2-����,3-����,4-��һ,5-�����շѱ�־��0-δ�շѻ���ʻ���,1-���շѻ����
    '   cllPatiData=������Ϣ���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiItem-������Ϣ����Ա(����ID,��������,���֤��,���")=cllPatiData(_����ID)
    '���Σ�
    '   strErrMsg=���ش�����Ϣ
    '   strNos=�漰�ĵ��ݺţ���ʽ��A001,A002,...
    '����:ִ�гɹ�����True��ִ��ʧ�ܷ���False
    Dim cllOrderList As Collection
    Dim cllExseBillList As Collection, cllExseBillItem As Collection
    Dim cllPatiBillItem_New As Collection, cllBillLists_New As Collection
    Dim strKey As String
    Dim cllPatiBillItem As Collection, cllBillLists As Collection
    Dim bln���ʱ� As Boolean
    
    On Error GoTo ErrHandler
    strNewBillCheckJson_Out = "": strNewBillJson_Out = "": strErrMsg = "": strNos = ""
    If cllOrderSendItem Is Nothing Then GetNewBillJson_Cis = True: Exit Function
    
    Set cllOrderList = cllOrderSendItem("OrderList")
    Set cllExseBillList = cllOrderSendItem("ExseBillList")
    
    '����ҽ���漰�ķ��õ��ݣ����鵥�ݼ�¼��
    Set cllPatiBillItem_New = New Collection
    Set cllBillLists_New = New Collection
    For Each cllExseBillItem In cllExseBillList
        strKey = "_" & cllExseBillItem("������Դ") & "_" & cllExseBillItem("��������") & "_" & cllExseBillItem("���ݺ�")
        For Each cllPatiBillItem In cllExseErrData
            Set cllBillLists = cllPatiBillItem("BillLists")
            If ExistsColObject(cllBillLists, strKey) Then
                If cllBillLists_New.Count = 0 Then
                    bln���ʱ� = (Val(Nvl(cllPatiBillItem("��������"))) = 3)
                    cllPatiBillItem_New.Add cllPatiBillItem("��������"), "��������"
                    cllPatiBillItem_New.Add cllPatiBillItem("������Դ"), "������Դ"
                    If bln���ʱ� = False Then
                        cllPatiBillItem_New.Add cllPatiBillItem("����ID"), "����ID"
                        cllPatiBillItem_New.Add cllPatiBillItem("��ҳID"), "��ҳID"
                        cllPatiBillItem_New.Add cllPatiBillItem("����"), "����"
                        cllPatiBillItem_New.Add cllPatiBillItem("�Ա���"), "�Ա���"
                        cllPatiBillItem_New.Add cllPatiBillItem("�Ա�"), "�Ա�"
                        cllPatiBillItem_New.Add cllPatiBillItem("����"), "����"
                        cllPatiBillItem_New.Add cllPatiBillItem("���˿���ID"), "���˿���ID"
                        cllPatiBillItem_New.Add cllPatiBillItem("���˲���ID"), "���˲���ID"
                    End If
                    cllPatiBillItem_New.Add cllBillLists_New, "BillLists"
                End If
                
                cllBillLists_New.Add cllBillLists(strKey), strKey
                Exit For
            End If
        Next
    Next
    If cllPatiBillItem_New.Count = 0 Then
        '�޷��õ���
        GetNewBillJson_Cis = True: Exit Function
    End If
    
    If GetNewBillJson_Exse(cllPatiBillItem_New, cllPatiData, strNewBillCheckJson_Out, strNewBillJson_Out, strErrMsg, strNos, cllOrderList) = False Then Exit Function
    
    GetNewBillJson_Cis = True
    Exit Function
ErrHandler:
    strErrMsg = err.Description
End Function

Private Function GetNewBillJson_Exse(ByVal cllPatiBillItem As Collection, ByVal cllPatiData As Collection, _
    ByRef strNewBillCheckJson_Out As String, ByRef strNewBillJson_Out As String, ByRef strErrMsg As String, ByRef strNos As String, _
    Optional ByVal cllOrderList As Collection) As Boolean
    'ִ���ٴ����쳣����ͬ��
    '��Σ�
    '   cllPatiBillItem-���˵��ݼ�¼����Ա(��������,������Դ,[����ID,��ҳID,����,�Ա���,�Ա�,����,���˿���ID,���˲���ID],BillLists)�����У��������е�Ԫ�ؼ��ʱ�ʱ��
    '           |-cllBillLists-������Ϣ��=cllPatiBillItem(BillLists)
    '               |-cllBillItem-������Ϣ����Ա(������Դ,NO,�շѱ�־,������,��������ID,������������,
    '                                                          ����ҽʦID,����ҽʦ,����Ա����,����Ա���,�Ǽ�ʱ��,DetailList)=cllBillLists(_������Դ_��������_���ݺ�)
    '                   |-cllDetailList-������ϸ��=cllBillItem(DetailList)
    '                       |-cllDetailItem-ÿ����ϸ���ݼ�����Ա([����ID,��ҳID,����,�Ա���,�Ա�,����,���˲���ID,���˿���ID],
    '                                 ����ID,���,�ⷿID,�Ƿ񱸻�����,����,����ID,Ӥ�����,ҽ��ID,����,����,�ۼ�,���۽��,ժҪ)�����У��������е�Ԫ�ؼ��ʱ�ʱ����
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ,4-��죻������Դ��1-����,2-סԺ��
    '                 �������ͣ�0�Ϳ�-��ͨ,1-����,2-����,3-����,4-��һ,5-�����շѱ�־��0-δ�շѻ���ʻ���,1-���շѻ����
    '   cllPatiData=������Ϣ���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiItem-������Ϣ����Ա(����ID,��������,���֤��,���")=cllPatiData(_����ID)
    '   cllOrderList-ҽ����Ϣ�б�
    '               |-cllOrderItem-ҽ����Ϣ����Ա(ҽ��ID,ҽ����Ч,������־,�Ƽ�����)=cllOrderList(_ҽ��ID)
    '���Σ�
    '   strNewBillCheckJson_Out=�µ��ݼ������JSON
    '   strNewBillJson_Out=�µ��ݱ�������JSON
    '   strErrMsg=���ش�����Ϣ
    '   strNos=�漰�ĵ��ݺţ���ʽ��A001,A002,...
    '����:ִ�гɹ�����True��ִ��ʧ�ܷ���False
    Dim strJson As String, bln���ʱ� As Boolean
    Dim cllBillLists As Collection, cllBillItem As Collection
    Dim cllDetailList As Collection, cllDetailItem As Collection
    Dim strBillListJson As String, strDetailListJson As String
    Dim rsTotal As adodb.Recordset, cllOrderItem As Collection
    Dim cllPati As Collection
    Dim strShowNO As String, strShow�������� As String, strShow���� As String
        
    On Error GoTo ErrHandler
    strNewBillCheckJson_Out = "": strNewBillJson_Out = "": strErrMsg = "": strNos = ""
    If cllPatiBillItem Is Nothing Then GetNewBillJson_Exse = True: Exit Function
    
    Set rsTotal = New adodb.Recordset
    rsTotal.Fields.Append "�ⷿID", adBigInt, , adFldIsNullable
    rsTotal.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsTotal.Fields.Append "����", adDouble, , adFldIsNullable
    rsTotal.Fields.Append "����", adDouble, , adFldIsNullable
    rsTotal.Fields.Append "�Ƿ񱸻�����", adInteger, , adFldIsNullable
    rsTotal.Fields.Append "����", adBigInt, , adFldIsNullable
    rsTotal.Fields.Append "����id", adBigInt, , adFldIsNullable
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    'Zl_ҩƷ�շ���¼_Newstuffbill
    '  --���ܣ���Ҫ���ڼ��ʣ������ۣ��� �շ�(������)������µĴ�����ҩ����¼
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --     billtype             N   1 ��������: 1 -�շѴ���  ;2- ���ʵ�����;3- ���ʱ���
    '  --     pati_source          N   1 ������Դ:1-����;2-סԺ;4-���
    '  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������½ڵ�--------------------------------------
    '  --     pati_id                    N   1 ����ID
    '  --     pati_pageid                N   1 ��ҳID
    '  --     pati_name                  C   1 ��������
    '  --     pati_sex_code              C   1 �Ա��ţ�������)
    '  --     pati_sex                   C   1 �Ա�
    '  --     pati_age                   C   1 ����
    '  --     pati_identity              C     ���
    '  --     pati_birthdate             C     ��������:yyyy-mm-dd hh:mi:ss
    '  --     pati_idcard                C     ���֤��
    '  --     pati_deptid                N   1 ���˿���ID
    '  --     pati_wardarea_id           N     ���˲���ID
    '  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������Ͻڵ�-----------------------------------------
    '  --     bill_list[]                      ���������б�[����]
    '  --        stuff_no                  C  1 NO
    '  --        charge_tag                N  1 �շѱ�־:0-δ�շѻ���ʻ���;1-���շѻ����
    '  --        fee_acnter                C    ������
    '  --        plcdept_id                C    ��������id��������)
    '  --        plcdept                   C    �����������ƣ�������)
    '  --        placer_id                 C    ����ҽʦid��������)
    '  --        placer                    C    ����ҽʦ��������)  ����
    '  --        apply_fee_category_code   C    ���뵥�ѱ����(ҽ�Ƹ��ʽ����)(������) ���ӣ�
    '  --        apply_fee_category_name   C    ���뵥�ѱ����ƣ�ҽ�Ƹ��ʽ���ƣ�(������) ���ӣ�
    '  --        operator_name             C  1 ����Ա����
    '  --        operator_code             C  1 ����Ա���
    '  --        create_time               C  1 �Ǽ�ʱ��:yyyy-mm-dd hh:mi:ss
    '  --        item_list[]                    ���������б�[����]
    '  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������½ڵ�----------------------------------------
    '  --           pati_id                 N  1 ����ID
    '  --           pati_pageid             N    ��ҳID
    '  --           pati_name               C  1 ��������
    '  --           pati_sex                C  1 �Ա�
    '  --           pati_age                C  1 ����
    '  --           pati_identity           C    ���
    '  --           pati_birthdate          C    ��������:yyyy-mm-dd hh:mi:ss
    '  --           pati_idcard             C    ���֤��
    '  --           pati_wardarea_id        N    ���˲���ID
    '  --           pati_deptid             N  1 ���˿���ID
    '  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������Ͻڵ�-----------------------------------------
    '  --           stuffdtl_id             N  1 ������ϸID
    '  --           serial_num              N  1 ���
    '  --           warehouse_id            N  1 �ⷿID
    '  --           is_bakstuff             N  1 �Ƿ񱸻�����:�и�ֵ���Ĳ���Ҫ���룬��0��ʾ�Ǹ�ֵ����ģʽ(��ɨ��ʱʹ��)
    '  --           bakstuff_batch             1 ������������
    '  --           stuff_id                N  1 ����ID
    '  --           baby_num                N    Ӥ�����
    '  ---------------------------���½ڵ�Ϊ��ѡ������ҽ����¼����-----------------------------------------------
    '  --           advice_id               N  0 ҽ��ID
    '  --           emergency_tag           N    ҽ����¼�еĽ�����־(0-��ͨ;1-����;2-��¼(��������Ч))
    '  --           effectivetime           N  0 ҽ����Ч
    '  --           freq_name               C  0 Ƶ������
    '  --           single                  N  0 ����
    '  ---------------------------���Ͻڵ�Ϊ��ѡ������ҽ����¼����-----------------------------------------------
    '  --           packages_num            N  1 ����
    '  --           outbound_num            N  1 ��������
    '  --           price                   N    �ۼ�
    '  --           warehouse_window        C  0 ���ϴ���
    '  --           memo                    C  0 ժҪ
    '  --           fee_source              N  0 ������Դ
    '  --           stuff_auto_send         N  0 �����Զ�����;0-���Զ�����;1-�Զ�����
    
    bln���ʱ� = (cllPatiBillItem("��������") = 3)
    Set cllBillLists = cllPatiBillItem("BillLists")
    
    strBillListJson = ""
    For Each cllBillItem In cllBillLists
        strDetailListJson = ""
        Set cllDetailList = cllBillItem("DetailList")
        
        For Each cllDetailItem In cllDetailList
            
            rsTotal.Filter = "�ⷿID=" & cllDetailItem("�ⷿID") & " And ����ID=" & cllDetailItem("����ID") & " And ����=" & Val(Nvl(cllDetailItem("����")))
            If rsTotal.EOF Then rsTotal.AddNew
            rsTotal!�ⷿID = cllDetailItem("�ⷿID")
            rsTotal!����ID = cllDetailItem("����ID")
            rsTotal!���� = Val(Nvl(rsTotal!����)) + IIf(cllDetailItem("����") = 0, 1, cllDetailItem("����")) * cllDetailItem("����")
            rsTotal!���� = Val(cllDetailItem("�ۼ�"))
            rsTotal!�Ƿ񱸻����� = Val(cllDetailItem("�Ƿ񱸻�����"))
            rsTotal!���� = Val(cllDetailItem("����"))
            rsTotal!����id = Val(cllDetailItem("����id"))
            
            Set cllOrderItem = Nothing
            If Not cllOrderList Is Nothing And Val(Nvl(cllDetailItem("ҽ��ID"))) <> 0 Then
                If ExistsColObject(cllOrderList, "_" & cllDetailItem("ҽ��ID")) Then Set cllOrderItem = cllOrderList("_" & cllDetailItem("ҽ��ID"))
            End If
            
            strJson = ""
            If bln���ʱ� Then
                strJson = strJson & "," & GetJsonNodeString("pati_id", cllDetailItem("����ID"), 1)
                strJson = strJson & "," & GetJsonNodeString("pati_pageid", cllDetailItem("��ҳID"), 1)
                strJson = strJson & "," & GetJsonNodeString("pati_name", cllDetailItem("����"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_sex_code", cllDetailItem("�Ա���"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_sex", cllDetailItem("�Ա�"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_age", cllDetailItem("����"), 0)
                
                Set cllPati = cllPatiData("_" & cllDetailItem("����ID")) '����ID,��������,���֤��,���
                strJson = strJson & "," & GetJsonNodeString("pati_identity", cllPati("���"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_birthdate", cllPati("��������"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_idcard", cllPati("���֤��"), 0)
                
                strJson = strJson & "," & GetJsonNodeString("pati_wardarea_id", cllDetailItem("���˲���ID"), 1)
                strJson = strJson & "," & GetJsonNodeString("pati_deptid", cllDetailItem("���˿���ID"), 1)
            End If
            strJson = strJson & "," & GetJsonNodeString("stuffdtl_id", cllDetailItem("����ID"), 1)
            strJson = strJson & "," & GetJsonNodeString("serial_num", cllDetailItem("���"), 1)
            strJson = strJson & "," & GetJsonNodeString("warehouse_id", cllDetailItem("�ⷿID"), 1)
            strJson = strJson & "," & GetJsonNodeString("is_bakstuff", cllDetailItem("�Ƿ񱸻�����"), 1)
            strJson = strJson & "," & GetJsonNodeString("bakstuff_batch", cllDetailItem("����"), 1)
            strJson = strJson & "," & GetJsonNodeString("stuff_id", cllDetailItem("����ID"), 1)
            strJson = strJson & "," & GetJsonNodeString("baby_num", cllDetailItem("Ӥ�����"), 1)
            strJson = strJson & "," & GetJsonNodeString("advice_id", cllDetailItem("ҽ��ID"), 1)
            If Not cllOrderItem Is Nothing Then
                strJson = strJson & "," & GetJsonNodeString("emergency_tag", cllOrderItem("������־"), 1)
                strJson = strJson & "," & GetJsonNodeString("effectivetime", cllOrderItem("ҽ����Ч"), 1)
                strJson = strJson & "," & GetJsonNodeString("freq_name", cllOrderItem("Ƶ������"), 0)
                strJson = strJson & "," & GetJsonNodeString("single", cllOrderItem("����"), 1)
            End If
            strJson = strJson & "," & GetJsonNodeString("packages_num", cllDetailItem("����"), 1)
            strJson = strJson & "," & GetJsonNodeString("outbound_num", cllDetailItem("����"), 1)
            strJson = strJson & "," & GetJsonNodeString("price", cllDetailItem("�ۼ�"), 1)
            strJson = strJson & "," & GetJsonNodeString("memo", cllDetailItem("ժҪ"), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_source", cllBillItem("������Դ"), 1)
            'strJson = strJson & "," & GetJsonNodeString("stuff_auto_send", cllDetailItem(""), 1) '�����Զ�����;0-���Զ�����;1-�Զ�����
            
            strDetailListJson = strDetailListJson & ",{" & Mid(strJson, 2) & "}"
        Next
        
        strJson = ""
        strJson = strJson & "," & GetJsonNodeString("stuff_no", cllBillItem("NO"), 0)
        strJson = strJson & "," & GetJsonNodeString("charge_tag", cllBillItem("�շѱ�־"), 1)
        strJson = strJson & "," & GetJsonNodeString("fee_acnter", cllBillItem("������"), 0)
        strJson = strJson & "," & GetJsonNodeString("plcdept_id", cllBillItem("��������ID"), 0)
        strJson = strJson & "," & GetJsonNodeString("plcdept", cllBillItem("������������"), 0)
        strJson = strJson & "," & GetJsonNodeString("placer_id", cllBillItem("����ҽʦID"), 0)
        strJson = strJson & "," & GetJsonNodeString("placer", cllBillItem("����ҽʦ"), 0)
        'strJson = strJson & "," & GetJsonNodeString("apply_fee_category_code", cllBillItem(""), 0)'���뵥�ѱ����(ҽ�Ƹ��ʽ����)(������)
        'strJson = strJson & "," & GetJsonNodeString("apply_fee_category_name", cllBillItem(""), 0)'���뵥�ѱ����ƣ�ҽ�Ƹ��ʽ���ƣ�(������)
        strJson = strJson & "," & GetJsonNodeString("operator_name", cllBillItem("����Ա����"), 0)
        strJson = strJson & "," & GetJsonNodeString("operator_code", cllBillItem("����Ա���"), 0)
        strJson = strJson & "," & GetJsonNodeString("create_time", cllBillItem("�Ǽ�ʱ��"), 0)
        strJson = strJson & ",""item_list"":[" & Mid(strDetailListJson, 2) & "]"
        strBillListJson = strBillListJson & ",{" & Mid(strJson, 2) & "}"
        
        If InStr("," & strNos & ",", "," & cllBillItem("NO") & ",") = 0 Then
            strNos = strNos & "," & cllBillItem("NO")
            
            strShowNO = cllBillItem("NO") & IIf(cllPatiBillItem("��������") = 1, "(�շ�)", "(����)")
            strShow���� = cllBillItem("������������")
            strShow�������� = cllPatiBillItem("����") & "(" & cllPatiBillItem("�Ա�") & "," & cllPatiBillItem("����") & ")"
            
            mstrShow = strShowNO & " " & strShow�������� & " " & strShow����
        End If
    Next
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("billtype", cllPatiBillItem("��������"), 1)
    strJson = strJson & "," & GetJsonNodeString("pati_source", cllPatiBillItem("������Դ"), 1)
    If bln���ʱ� = False Then
        strJson = strJson & "," & GetJsonNodeString("pati_id", cllPatiBillItem("����ID"), 1)
        strJson = strJson & "," & GetJsonNodeString("pati_pageid", cllPatiBillItem("��ҳID"), 1)
        strJson = strJson & "," & GetJsonNodeString("pati_name", cllPatiBillItem("����"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_sex_code", cllPatiBillItem("�Ա���"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_sex", cllPatiBillItem("�Ա�"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_age", cllPatiBillItem("����"), 0)
        
        Set cllPati = cllPatiData("_" & cllPatiBillItem("����ID")) '����ID,��������,���֤��,���
        strJson = strJson & "," & GetJsonNodeString("pati_identity", cllPati("���"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_birthdate", cllPati("��������"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_idcard", cllPati("���֤��"), 0)
        
        strJson = strJson & "," & GetJsonNodeString("pati_deptid", cllPatiBillItem("���˿���ID"), 1)
        strJson = strJson & "," & GetJsonNodeString("pati_wardarea_id", cllPatiBillItem("���˲���ID"), 1)
    End If
    strJson = strJson & ",""bill_list"":[" & Mid(strBillListJson, 2) & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If GetNewBillCheckJson(rsTotal, strNewBillCheckJson_Out) = False Then Exit Function
    
    strNewBillJson_Out = strJson
    GetNewBillJson_Exse = True
    Exit Function
ErrHandler:
    strErrMsg = err.Description
End Function

Private Function GetNewBillCheckJson(ByVal rsTotal As adodb.Recordset, ByRef strCheckJson_Out As String) As Boolean
    '����:��ȡ����ҩƷ�������������Json��δ�
    '���:
    '   rsTotal-��ǰ�Ļ��ܼ�¼��(����ID,�ⷿID,����,����)
    '����:
    '����:����Json��
    Dim strJson As String, strListJson As String
    
    strCheckJson_Out = ""
    If rsTotal Is Nothing Then GetNewBillCheckJson = True: Exit Function

    'Zl_�������۳���_Check
    '  --���      json
    '  --input     ����������Ҫ�����Ĵ������м��
    '  --  fee_list      �շ���ϸ��Ϣ��֧�ֶ����[����]
    '  --    stuff_id  N 1 ����id
    '  --    send_num  N 1 ��������
    '  --    warehouse_id  N 1 �ⷿid
    '  --    price           N       1       �ۼ�
    '  --    is_bakstuff N   �Ƿ񱸻�����:�и�ֵ���Ĳ���Ҫ���룬��0��ʾ�Ǹ�ֵ����ģʽ(��ɨ��ʱʹ��)
    '  --    bakstuff_batch  N   ������������
    '  --    rcpdtl_id    N  1  ����id��0���-û�д���ʱ���ԣ�>0����ʱ����Ƿ��Ѵ�����ͬ�ķ���ID�շ���¼
    With rsTotal
        .Filter = ""
        Do While Not .EOF
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("stuff_id", Val(Nvl(!����ID)), 1)
            strJson = strJson & "," & GetJsonNodeString("send_num", Val(Nvl(!����)), 1)
            strJson = strJson & "," & GetJsonNodeString("warehouse_id", Val(Nvl(!�ⷿID)), 1)
            strJson = strJson & "," & GetJsonNodeString("price", Val(Nvl(!����)), 1)
            strJson = strJson & "," & GetJsonNodeString("is_bakstuff", Val(Nvl(!�Ƿ񱸻�����)), 1)
            strJson = strJson & "," & GetJsonNodeString("bakstuff_batch", Val(Nvl(!����)), 1)
            strJson = strJson & "," & GetJsonNodeString("rcpdtl_id", Val(Nvl(!����id)), 1)
            strListJson = strListJson & ",{" & strJson & "}"
            .MoveNext
        Loop
    End With
    If strListJson = "" Then GetNewBillCheckJson = True: Exit Function
    
    strCheckJson_Out = "{""input"":{""fee_list"":[" & Mid(strListJson, 2) & "]}}"
    GetNewBillCheckJson = True
End Function

Private Function GetPatiData(ByVal cllExseErrData As Collection, ByRef cllPatiData As Collection, ByRef strErrMsg As String) As Boolean
    '��ȡ��������
    '��Σ�
    '   cllExseErrData=������ͬ���쳣���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiBillItem-���˵��ݼ�¼����Ա(��������,������Դ,[����ID,��ҳID,����,�Ա���,�Ա�,����,���˿���ID,���˲���ID],BillLists)�����У��������е�Ԫ�ؼ��ʱ�ʱ��
    '           |-cllBillLists-������Ϣ��=cllPatiBillItem(BillLists)
    '               |-cllBillItem-������Ϣ����Ա(������Դ,NO,�շѱ�־,������,��������ID,������������,
    '                                                          ����ҽʦID,����ҽʦ,����Ա����,����Ա���,�Ǽ�ʱ��,DetailList)=cllBillLists(_������Դ_��������_���ݺ�)
    '                   |-cllDetailList-������ϸ��=cllBillItem(DetailList)
    '                       |-cllDetailItem-ÿ����ϸ���ݼ�����Ա([����ID,��ҳID,����,�Ա���,�Ա�,����,���˲���ID,���˿���ID],
    '                                 ����ID,���,�ⷿID,�Ƿ񱸻�����,����,����ID,Ӥ�����,ҽ��ID,����,����,�ۼ�,���۽��,ժҪ)�����У��������е�Ԫ�ؼ��ʱ�ʱ����
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ,4-��죻������Դ��1-����,2-סԺ���շѱ�־��0-δ�շѻ���ʻ���,1-���շѻ����
    '���Σ�
    '   cllPatiData=������Ϣ���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiItem-������Ϣ����Ա(����ID,����,�����,סԺ��,��������,���֤��,���")=cllPatiData(_����ID)
    '   strErrMsg=��ΧֵΪ2ʱ�����ش�����Ϣ
    '���أ��ɹ�����True��ʧ�ܷ���False
    Dim bln���ʱ� As Boolean, strPatiIDs As String, cllItem As Collection
    Dim cllPatiBillItem As Collection, cllBillLists As Collection
    Dim cllDetailList As Collection, cllDetailItem As Collection
    Dim cllPatiOut As Collection, cllPati As Collection
    Dim p As Long, i As Long, j As Long
    Dim StrJson_In As String
    
    On Error GoTo ErrHandler
    Set cllPatiData = New Collection
    strErrMsg = ""
    
    If cllExseErrData Is Nothing Then GetPatiData = True: Exit Function
    For p = 1 To cllExseErrData.Count
        Set cllPatiBillItem = cllExseErrData(p)
        bln���ʱ� = (Val(cllPatiBillItem("��������")) = 3)
        
        If bln���ʱ� = False Then
            If InStr("," & strPatiIDs & ",", "," & cllPatiBillItem("����ID") & ",") = 0 Then
                strPatiIDs = strPatiIDs & "," & cllPatiBillItem("����ID")
            End If
        Else
            Set cllBillLists = cllPatiBillItem("BillLists")
            For i = 1 To cllBillLists.Count
                Set cllDetailList = cllBillLists(i)("DetailList")
                For j = 1 To cllDetailList.Count
                    Set cllDetailItem = cllDetailList(j)
                    If InStr("," & strPatiIDs & ",", "," & cllDetailItem("����ID") & ",") = 0 Then
                        strPatiIDs = strPatiIDs & "," & cllDetailItem("����ID")
                    End If
                Next
            Next
        End If
    Next
    
    'Zl_Patisvr_Getpatiinfo
    '  --����:��ȡ������Ϣ
    '  --��Σ�Json_In:��ʽ
    '  --    input
    '  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
    '  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;2-����
    '  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
    '  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
    '  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
    '  --      query_immune      N 1 �Ƿ����������:1-����;0-������
    '  --      query_insurance_pwd C  �Ƿ����ҽ������:1-����;0-������
    '  --      query_cons_list   C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
    '  --        pati_ids        C   ����IDs:����ö���
    '  --        pati_name       C   ����:���Դ�%�ֺű������ƥ��
    '  --        outpatient_num  C   �����
    '  --        inpatient_num   C   סԺ��
    '  --        pati_idcard     C   ���֤��
    '  --        contacts_idcard C   ��ϵ�����֤��
    '  --        cardtype_id     N   ҽ�ƿ����ID
    '  --        medc_card_name  N   ҽ�ƿ�����
    '  --        card_no         C   ����
    '  --        qrcode          C   ��ά��
    '  --        iccard_no       C   Ic����
    '  --        visit_card      C   ���￨��
    '  --        insurance_num   C   ҽ����
    '  --        qrspt_statu     C   ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
    '  --        phone_number    C   �ֻ���
    '  --        pati_bed        C   ��ǰ����
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 3, 1)
    StrJson_In = StrJson_In & "," & """query_cons_list"":{""qrspt_statu"":2,""pati_ids"":""" & Mid(strPatiIDs, 2) & """}"
    StrJson_In = "{""input"":{" & StrJson_In & "}}"

    If mobjServiceCall.CallService("Zl_Patisvr_Getpatiinfo", StrJson_In, , , , False, , , , True) = False Then
        strErrMsg = "���÷����ѯ������Ϣʧ�ܣ�"
        Exit Function
    End If

    Set cllPatiOut = mobjServiceCall.GetJsonListValue("output.pati_list", "pati_id")
    If cllPatiOut Is Nothing Then Exit Function
    
    For i = 1 To cllPatiOut.Count
        '--    pati_list[]                 ������Ϣ�б�
        '--      pati_id             N   1   ����id
        '--      pati_name           C   1   ����
        '--      outpatient_num      C   1   �����
        '--      inpatient_num       C   1   סԺ��
        '--      pati_birthdate      C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
        '  --    pati_idcard         C   1   ���֤��
        '--      pati_identity       C   1   ���
        Set cllItem = cllPatiOut(i)
        Set cllPati = New Collection
        cllPati.Add cllItem("_pati_id"), "����ID"
        cllPati.Add cllItem("_pati_name"), "����"
        cllPati.Add cllItem("_outpatient_num"), "�����"
        cllPati.Add cllItem("_inpatient_num"), "סԺ��"
        cllPati.Add cllItem("_pati_birthdate"), "��������"
        cllPati.Add cllItem("_pati_idcard"), "���֤��"
        cllPati.Add cllItem("_pati_identity"), "���"
        cllPatiData.Add cllPati, "_" & cllItem("_pati_id")
    Next
    GetPatiData = True
    Exit Function
ErrHandler:
    strErrMsg = err.Description
End Function

Private Function GetExseSyncErrData(ByVal strPatiIDs As String, ByVal cllCisErrData As Collection, _
    ByRef cllExseErrData As Collection, ByRef strErrMsg As String) As Integer
    '��ȡҽ���������ݼ�����ͬ���쳣����
    '��Σ�
    '   strPatiIDs=����ID,�����Ӣ�ĵĶ��ŷָ�
    '   cllCisErrData-�ٴ���ͬ���쳣���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllOrderSendItem-����ҽ�����ͼ�¼����Ա(����ID,��ҳID,�Һ�ID,�Һŵ���,���ͺ�,OrderList)
    '           |-cllOrderList-ҽ����Ϣ�б�=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-ҽ����Ϣ����Ա(ҽ��ID,ҽ����Ч,������־,�Ƽ�����)=cllOrderList(_ҽ��ID)
    '           |-cllExseBillList-���õ����б�=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-���õ�����Ϣ����Ա(������Դ,��������,���ݺ�)=cllExseBillList(_������Դ_��������_���ݺ�)
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ
    '���Σ�
    '   cllExseErrData=������ͬ���쳣���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiBillItem-���˵��ݼ�¼����Ա(��������,������Դ,[����ID,��ҳID,����,�Ա���,�Ա�,����,���˿���ID,���˲���ID],BillLists)�����У��������е�Ԫ�ؼ��ʱ�ʱ��
    '           |-cllBillLists-������Ϣ��=cllPatiBillItem(BillLists)
    '               |-cllBillItem-������Ϣ����Ա(������Դ,NO,�շѱ�־,������,��������ID,������������,
    '                                                          ����ҽʦID,����ҽʦ,����Ա����,����Ա���,�Ǽ�ʱ��,DetailList)=cllBillLists(_������Դ_��������_���ݺ�)
    '                   |-cllDetailList-������ϸ��=cllBillItem(DetailList)
    '                       |-cllDetailItem-ÿ����ϸ���ݼ�����Ա([����ID,��ҳID,����,�Ա���,�Ա�,����,���˲���ID,���˿���ID],
    '                                 ����ID,���,�ⷿID,�Ƿ񱸻�����,����,����ID,Ӥ�����,ҽ��ID,����,����,�ۼ�,���۽��,ժҪ)�����У��������е�Ԫ�ؼ��ʱ�ʱ����
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ,4-��죻������Դ��1-����,2-סԺ��
    '                 �������ͣ�0�Ϳ�-��ͨ,1-����,2-����,3-����,4-��һ,5-�����շѱ�־��0-δ�շѻ���ʻ���,1-���շѻ����
    '   strErrMsg=��ΧֵΪ2ʱ�����ش�����Ϣ
    '���أ�0-����δͬ���ĵ��ݣ�1-������δͬ���ĵ��ݣ�2-��������
    Dim StrJson_In As String, strJson_List As String, strJsonItem As String, strJson_PatiList As String
    Dim cllExseBillList As Collection, cllItem As Collection
    Dim p As Long, i As Long, j As Long
    Dim cllOutList As Collection, cllBill_Out As Collection, cllDetail_Out As Collection
    Dim cllPatiBillItem As Collection, cllBillLists As Collection, cllBillItem As Collection
    Dim cllDetailList As Collection, cllDetailItem As Collection
    Dim bln���ʱ� As Boolean, varPatiIDs As Variant
    Dim strKey As String, byt�������� As Byte
    
    On Error GoTo ErrHandler
    Set cllExseErrData = New Collection
    strErrMsg = ""
    
    If strPatiIDs = "" Then GetExseSyncErrData = 1: Exit Function
    'Zl_Exsesvr_Getstufferrdata
    '  --���ܣ����ݲ���ID��ҽ����Ϣ���ز��˷�����Ϣ
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --    pati_list[]�����б�
    '  --       pati_id                    N 1 ����id
    '  --       bill_list[]                ���õ��ݺ��б����Բ���������ʱ��ʾ��ȡ������ͬ���쳣������
    '  --         fee_source               N 0 ������Դ��1-���2-סԺ
    '  --         fee_billtype             N 0 ���õ������ͣ�1-�շѴ�����2-���ʵ�����
    '  --         fee_no                   C 0 ���õ��ݺ�
    strJson_PatiList = ""
    varPatiIDs = Split(strPatiIDs, ",")
    For p = 0 To UBound(varPatiIDs)
        strJson_List = ""
        If Not cllCisErrData Is Nothing Then
            For i = 1 To cllCisErrData.Count
                Set cllItem = cllCisErrData(i)
                
                If Val(Nvl(cllItem("����ID"))) = varPatiIDs(p) Then
                    Set cllExseBillList = cllItem("ExseBillList")
                    For j = 1 To cllExseBillList.Count
                        Set cllItem = cllExseBillList(j)
                        strJsonItem = ""
                        strJsonItem = strJsonItem & "" & GetJsonNodeString("fee_source", cllItem("������Դ"), 1)
                        strJsonItem = strJsonItem & "," & GetJsonNodeString("fee_billtype", cllItem("��������"), 1)
                        strJsonItem = strJsonItem & "," & GetJsonNodeString("fee_no", cllItem("���ݺ�"), 0)
                        strJson_List = strJson_List & ",{" & strJsonItem & "}"
                    Next
                End If
            Next
        End If
        
        strJsonItem = ""
        strJsonItem = strJsonItem & "" & GetJsonNodeString("pati_id", varPatiIDs(p), 1)
        If strJson_List <> "" Then
            strJsonItem = strJsonItem & ",""bill_list"":[" & Mid(strJson_List, 2) & "]"
        End If
        strJson_PatiList = strJson_PatiList & ",{" & strJsonItem & "}"
    Next
    StrJson_In = "{""input"":{""pati_list"":[" & Mid(strJson_PatiList, 2) & "]}}"
    
    If mobjServiceCall.CallService("Zl_Exsesvr_Getstufferrdata", StrJson_In, , , , False, , , , True) = False Then
        strErrMsg = "���÷��÷����ѯδ��������ʧ�ܣ�"
        GetExseSyncErrData = 2: Exit Function
    End If
    
    Set cllOutList = mobjServiceCall.GetJsonListValue("output.pati_bill_list")
    If cllOutList Is Nothing Then GetExseSyncErrData = 1: Exit Function
    If cllOutList.Count = 0 Then GetExseSyncErrData = 1: Exit Function

    '   cllExseErrData=������ͬ���쳣���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllPatiBillItem-���˵��ݼ�¼����Ա(��������,������Դ,[����ID,��ҳID,����,�Ա���,�Ա�,����,���˿���ID,���˲���ID],BillLists)�����У��������е�Ԫ�ؼ��ʱ�ʱ��
    '           |-cllBillLists-������Ϣ��=cllPatiBillItem(BillLists)
    '               |-cllBillItem-������Ϣ����Ա(������Դ,NO,�շѱ�־,������,��������ID,������������,
    '                                                          ����ҽʦID,����ҽʦ,����Ա����,����Ա���,�Ǽ�ʱ��,DetailList)=cllBillLists(_������Դ_��������_���ݺ�)
    '                   |-cllDetailList-������ϸ��=cllBillItem(DetailList)
    '                       |-cllDetailItem-ÿ����ϸ���ݼ�����Ա([����ID,��ҳID,����,�Ա���,�Ա�,����,���˲���ID,���˿���ID],
    '                                 ����ID,���,�ⷿID,�Ƿ񱸻�����,����,����ID,Ӥ�����,ҽ��ID,����,����,�ۼ�,���۽��,ժҪ)�����У��������е�Ԫ�ؼ��ʱ�ʱ����
    Set cllExseErrData = New Collection
    For p = 1 To cllOutList.Count
        '  --    pati_bill_list[]
        '  --       billtype                   N   1 ��������: 1 -�շѴ���  ;2- ���ʵ�����;3- ���ʱ���
        '  --       pati_source                N   1 ������Դ:1-����;2-סԺ;4-���
        '  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������½ڵ�--------------------------------------
        '  --       pati_id                    N   1 ����ID
        '  --       pati_pageid                N   1 ��ҳID
        '  --       pati_name                  C   1 ��������
        '  --       pati_sex_code              C   1 �Ա��ţ�������)
        '  --       pati_sex                   C   1 �Ա�
        '  --       pati_age                   C   1 ����
        '  --       pati_deptid                N   1 ���˿���ID
        '  --       pati_wardarea_id           N     ���˲���ID
        '  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������Ͻڵ�-----------------------------------------
        '  --       bill_list[]                      ���������б�[����]
        Set cllBillLists = New Collection
        
        Set cllItem = cllOutList(p)
        Set cllPatiBillItem = New Collection
        byt�������� = Val(Nvl(cllItem("_billtype")))
        bln���ʱ� = (byt�������� = 3)
        cllPatiBillItem.Add cllItem("_billtype"), "��������"
        cllPatiBillItem.Add cllItem("_pati_source"), "������Դ"
        If bln���ʱ� = False Then
            cllPatiBillItem.Add cllItem("_pati_id"), "����ID"
            cllPatiBillItem.Add cllItem("_pati_pageid"), "��ҳID"
            cllPatiBillItem.Add cllItem("_pati_name"), "����"
            cllPatiBillItem.Add cllItem("_pati_sex_code"), "�Ա���"
            cllPatiBillItem.Add cllItem("_pati_sex"), "�Ա�"
            cllPatiBillItem.Add cllItem("_pati_age"), "����"
            cllPatiBillItem.Add cllItem("_pati_deptid"), "���˿���ID"
            cllPatiBillItem.Add cllItem("_pati_wardarea_id"), "���˲���ID"
        End If
        cllPatiBillItem.Add cllBillLists, "BillLists"
        cllExseErrData.Add cllPatiBillItem
        
        Set cllBill_Out = mobjServiceCall.GetJsonListValue("output.pati_bill_list[" & p - 1 & "].bill_list")
        For i = 1 To cllBill_Out.Count
            '  --       bill_list[]                      ���������б�[����]
            '  --         fee_source                N  0 ������Դ
            '  --         stuff_no                  C  1 NO
            '  --         charge_tag                N  1 �շѱ�־:0-δ�շѻ���ʻ���;1-���շѻ����
            '  --         fee_acnter                C  0 ������
            '  --         plcdept_id                C  0 ��������id��������)
            '  --         plcdept                   C  0 �����������ƣ�������)
            '  --         placer_id                 C  0 ����ҽʦid��������)
            '  --         placer                    C  0 ����ҽʦ��������) ����
            '  --         operator_name             C  1 ����Ա����
            '  --         operator_code             C  1 ����Ա���
            '  --         create_time               C  1 �Ǽ�ʱ��:yyyy-mm-dd hh:mi:ss
            '  --         item_list[]                    ���������б�[����]
            Set cllDetailList = New Collection
            
            Set cllItem = cllBill_Out(i)
            strKey = "_" & cllItem("_fee_source") & "_" & byt�������� & "_" & cllItem("_stuff_no")
            Set cllBillItem = New Collection
            cllBillItem.Add cllItem("_fee_source"), "������Դ"
            cllBillItem.Add cllItem("_stuff_no"), "NO"
            cllBillItem.Add cllItem("_charge_tag"), "�շѱ�־"
            cllBillItem.Add cllItem("_fee_acnter"), "������"
            cllBillItem.Add cllItem("_plcdept_id"), "��������ID"
            cllBillItem.Add cllItem("_plcdept"), "������������"
            cllBillItem.Add cllItem("_placer_id"), "����ҽʦID"
            cllBillItem.Add cllItem("_placer"), "����ҽʦ"
            cllBillItem.Add cllItem("_operator_name"), "����Ա����"
            cllBillItem.Add cllItem("_operator_code"), "����Ա���"
            cllBillItem.Add cllItem("_create_time"), "�Ǽ�ʱ��"
            cllBillItem.Add cllDetailList, "DetailList"
            cllBillLists.Add cllBillItem, strKey
            
            Set cllDetail_Out = mobjServiceCall.GetJsonListValue("output.pati_bill_list[" & p - 1 & "].bill_list[" & i - 1 & "].item_list")
            For j = 1 To cllDetail_Out.Count
                '  --         item_list[]                    ���������б�[����]
                '  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������½ڵ�----------------------------------------
                '  --           pati_id                 N  1 ����ID
                '  --           pati_pageid             N  0 ��ҳID
                '  --           pati_name               C  1 ��������
                '  --           pati_sex_code           C  1 �Ա��ţ�������)
                '  --           pati_sex                C  1 �Ա�
                '  --           pati_age                C  1 ����
                '  --           pati_wardarea_id        N    ���˲���ID
                '  --           pati_deptid             N  1 ���˿���ID
                '  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������Ͻڵ�-----------------------------------------
                '  --           stuffdtl_id             N  1 ������ϸID(Ŀǰ������Ƿ���id)
                '  --           serial_num              N  1 ���:(���(�����洢)����ź���ţ�1��2��3��3��3��4��)
                '  --           warehouse_id            N  1 �ⷿID
                '  --           is_bakstuff             N  1 �Ƿ񱸻�����:�и�ֵ���Ĳ���Ҫ���룬��0��ʾ�Ǹ�ֵ����ģʽ(��ɨ��ʱʹ��)
                '  --           bakstuff_batch          N  1 ������������
                '  --           stuff_id                N  1 ����ID
                '  --           baby_num                N  0 Ӥ�����
                '  --           advice_id               N  0 ҽ��ID
                '  --           packages_num            N  1 ����
                '  --           outbound_num            N  1 ��������
                '  --           price                   N  0 �ۼ�
                '  --           money                   N  0 ���۽��(������)
                '  --           memo                    C  0 ժҪ
                Set cllItem = cllDetail_Out(j)
                Set cllDetailItem = New Collection
                If bln���ʱ� Then
                    cllDetailItem.Add cllItem("_pati_id"), "����ID"
                    cllDetailItem.Add cllItem("_pati_pageid"), "��ҳID"
                    cllDetailItem.Add cllItem("_pati_name"), "����"
                    cllDetailItem.Add cllItem("_pati_sex_code"), "�Ա���"
                    cllDetailItem.Add cllItem("_pati_sex"), "�Ա�"
                    cllDetailItem.Add cllItem("_pati_age"), "����"
                    cllDetailItem.Add cllItem("_pati_wardarea_id"), "���˲���ID"
                    cllDetailItem.Add cllItem("_pati_deptid"), "���˿���ID"
                End If
                cllDetailItem.Add cllItem("_stuffdtl_id"), "����ID"
                cllDetailItem.Add cllItem("_serial_num"), "���"
                cllDetailItem.Add cllItem("_warehouse_id"), "�ⷿID"
                cllDetailItem.Add cllItem("_is_bakstuff"), "�Ƿ񱸻�����"
                cllDetailItem.Add cllItem("_bakstuff_batch"), "����"
                cllDetailItem.Add cllItem("_stuff_id"), "����ID"
                cllDetailItem.Add cllItem("_baby_num"), "Ӥ�����"
                cllDetailItem.Add cllItem("_advice_id"), "ҽ��ID"
                cllDetailItem.Add cllItem("_packages_num"), "����"
                cllDetailItem.Add cllItem("_outbound_num"), "����"
                cllDetailItem.Add cllItem("_price"), "�ۼ�"
                cllDetailItem.Add cllItem("_money"), "���۽��"
                cllDetailItem.Add cllItem("_memo"), "ժҪ"
                cllDetailList.Add cllDetailItem
            Next
        Next
    Next
 
    Exit Function
ErrHandler:
    GetExseSyncErrData = 2
    strErrMsg = err.Description
End Function

Private Function GetCisSyncErrData(ByVal strPatiIDs As String, ByRef cllCisErrData As Collection, ByRef strErrMsg As String) As Integer
    '��ȡ�ٴ���ͬ���쳣����
    '��Σ�
    '   strPatiIDs=����ID,�����Ӣ�ĵĶ��ŷָ�
    '���Σ�
    '   cllCisErrData-�ٴ���ͬ���쳣���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllOrderSendItem-����ҽ�����ͼ�¼����Ա(����ID,��ҳID,�Һ�ID,�Һŵ���,���ͺ�,OrderList)
    '           |-cllOrderList-ҽ����Ϣ�б�=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-ҽ����Ϣ����Ա(ҽ��ID,ҽ����Ч,������־,�Ƽ�����)=cllOrderList(_ҽ��ID)
    '           |-cllExseBillList-���õ����б�=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-���õ�����Ϣ����Ա(������Դ,��������,���ݺ�)=cllExseBillList(_������Դ_��������_���ݺ�)
    '       ���У��������ͣ�1-�շѵ�,2-���ʵ�,3-���ʱ�������Դ��1-����,2-סԺ
    '   strErrMsg=��ΧֵΪ2ʱ�����ش�����Ϣ
    '���أ�0-����δͬ���ĵ��ݣ�1-������δͬ���ĵ��ݣ�2-��������
    Dim StrJson_In As String, strKey As String
    Dim i As Long, j As Long
    Dim cllOutList As Collection, cllOrder_Out As Collection
    Dim cllOrderSendItem As Collection, cllItem As Collection
    Dim cllOrderList As Collection, cllOrderItem As Collection
    Dim cllExseBillList As Collection, cllExseBillItem As Collection
    
    On Error GoTo ErrHandler
    Set cllCisErrData = New Collection
    strErrMsg = ""

    If strPatiIDs = "" Then GetCisSyncErrData = 1: Exit Function
    'Zl_Cissvr_Getstufferrdata
'  --���ܣ��ٴ�ҽ������������������ͬ��
'  --��Σ�Json_In:��ʽ
'  --  input
'  --      pati_ids                        C 1 ����ids����ƴ��
'  --����: Json_Out,��ʽ����
'  --   output:
'  --     code: 1,
'  --     message: �ɹ�,
'  --     data[]
'  --         pati_id                      N 1 ����id
'  --         pati_pageid                  N 0 ��ҳid��סԺ���˴��룬���ﴫ0
'  --         rgst_id                      N 0 �Һ�id�����ﲡ�˴��룬סԺ���˴���
'  --         rgst_no                      C 0 �Һŵ���
'  --         send_no                      N 1 ���ͺ�
'  --         order_list[]ҽ����Ϣ�б�
'  --             advice_id                N 1 ҽ��id
'  --             effectivetime            N 1 ҽ����Ч
'  --             emergency_tag            N 1 ������־
'  --             denominated              N 1 �Ƽ�����
'  --             fee_source               N 0 ������Դ��1-���2-סԺ
'  --             fee_billtype             N 0 ���õ������ͣ�1-�շѴ�����2-���ʵ�����
'  --             fee_no                   C 0 ���õ��ݺ�
'  --             freq_name                C 0 Ƶ������
'  --             single                   N 0 ����
'  -----------------------------------------------------------------------------------
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_ids", strPatiIDs, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
 
    If mobjServiceCall.CallService("Zl_Cissvr_Getstufferrdata", StrJson_In, , , , False, , , , True) = False Then
        strErrMsg = "����ҽ�������ѯδ��������ʧ�ܣ�"
        GetCisSyncErrData = 2: Exit Function
    End If
    
    Set cllOutList = mobjServiceCall.GetJsonListValue("output.pati_bill_list")
    If cllOutList Is Nothing Then GetCisSyncErrData = 1: Exit Function
    If cllOutList.Count = 0 Then GetCisSyncErrData = 1: Exit Function
        
    '   cllCisErrData-�ٴ���ͬ���쳣���ݣ�˵���������еľ�Ϊ����Keyֵ
    '       |-cllOrderSendItem-����ҽ�����ͼ�¼����Ա(����ID,��ҳID,�Һ�ID,�Һŵ���,���ͺ�,OrderList)
    '           |-cllOrderList-ҽ����Ϣ�б�=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-ҽ����Ϣ����Ա(ҽ��ID,ҽ����Ч,������־,�Ƽ�����)=cllOrderList(_ҽ��ID)
    '           |-cllExseBillList-���õ����б�=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-���õ�����Ϣ����Ա(������Դ,��������,���ݺ�)=cllExseBillList(_������Դ_��������_���ݺ�)
    Set cllCisErrData = New Collection
    For i = 1 To cllOutList.Count
        '  --     pati_bill_list[]
        '  --         pati_id                      N 1 ����id
        '  --         pati_pageid                  N 0 ��ҳid��סԺ���˴��룬���ﴫ0
        '  --         rgst_id                      N 0 �Һ�id�����ﲡ�˴��룬סԺ���˴���
        '  --         rgst_no                      C 0 �Һŵ���
        '  --         send_no                      N 1 ���ͺ�
        '  --         order_list[]ҽ����Ϣ�б�
        Set cllOrderList = New Collection
        Set cllExseBillList = New Collection
        
        Set cllItem = cllOutList(i)
        Set cllOrderSendItem = New Collection
        cllOrderSendItem.Add cllItem("_pati_id"), "����ID"
        cllOrderSendItem.Add cllItem("_pati_pageid"), "��ҳID"
        cllOrderSendItem.Add cllItem("_rgst_id"), "�Һ�ID"
        cllOrderSendItem.Add cllItem("_rgst_no"), "�Һŵ���"
        cllOrderSendItem.Add cllItem("_send_no"), "���ͺ�"
        cllOrderSendItem.Add cllOrderList, "OrderList"
        cllOrderSendItem.Add cllExseBillList, "ExseBillList"
        cllCisErrData.Add cllOrderSendItem
        
        Set cllOrder_Out = mobjServiceCall.GetJsonListValue("output.pati_bill_list[" & i - 1 & "].order_list")
        For j = 1 To cllOrder_Out.Count
            '  --         order_list[]ҽ����Ϣ�б�
            '  --             advice_id                N 1 ҽ��id
            '  --             effectivetime            N 1 ҽ����Ч
            '  --             emergency_tag            N 1 ������־
            '  --             denominated              N 1 �Ƽ�����
            '  --             fee_source               N 0 ������Դ��1-���2-סԺ
            '  --             fee_billtype             N 0 ���õ������ͣ�1-�շѴ�����2-���ʵ�����
            '  --             fee_no                   C 0 ���õ��ݺ�
            '  --             freq_name                C 0 Ƶ������
            '  --             single                   N 0 ����
            Set cllItem = cllOrder_Out(j)
             
            '����ҽ����Ϣ�б���ͬ��ֻ��һ��
            strKey = "_" & cllItem("_advice_id")
            If ExistsColObject(cllOrderList, strKey) = False Then
                Set cllOrderItem = New Collection
                cllOrderItem.Add cllItem("_advice_id"), "ҽ��ID"
                cllOrderItem.Add cllItem("_effectivetime"), "ҽ����Ч"
                cllOrderItem.Add cllItem("_emergency_tag"), "������־"
                cllOrderItem.Add cllItem("_denominated"), "�Ƽ�����"
                cllOrderItem.Add cllItem("_freq_name"), "Ƶ������"
                cllOrderItem.Add cllItem("_single"), "����"
                cllOrderList.Add cllOrderItem, strKey
            End If
            
            '������õ�����Ϣ�б���ͬ��ֻ��һ��
            strKey = "_" & cllItem("_fee_source") & "_" & cllItem("_fee_billtype") & "_" & cllItem("_fee_no")
            If ExistsColObject(cllExseBillList, strKey) = False Then
                Set cllExseBillItem = New Collection
                '  --             fee_source               N 0 ������Դ��1-���2-סԺ
                '  --             fee_billtype             N 0 ���õ������ͣ�1-�շѴ�����2-���ʵ�����
                '  --             fee_no                   C 0 ���õ��ݺ�
                cllExseBillItem.Add cllItem("_fee_source"), "������Դ"
                cllExseBillItem.Add cllItem("_fee_billtype"), "��������"
                cllExseBillItem.Add cllItem("_fee_no"), "���ݺ�"
                cllExseBillList.Add cllExseBillItem, strKey
            End If
        Next
    Next
    
    Exit Function
ErrHandler:
    GetCisSyncErrData = 2
    strErrMsg = err.Description
End Function

Private Sub txtPatiId_Change()
    txtPatiId.Tag = ""
End Sub
