VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemSelectMulit 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   Icon            =   "frmItemSelectMulit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6300
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2145
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3210
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   540
      Width           =   45
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6300
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3900
      Width           =   6300
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4035
         TabIndex        =   4
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5265
         TabIndex        =   3
         Top             =   105
         Width           =   1100
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6300
      TabIndex        =   0
      Top             =   0
      Width           =   6300
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ��һ����Ŀ,Ȼ����ȷ��"
         Height          =   180
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   2430
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3240
      Left            =   2205
      TabIndex        =   5
      Top             =   555
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   5715
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   3240
      Left            =   15
      TabIndex        =   6
      Top             =   540
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   5715
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   4725
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelectMulit.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   2400
      ScaleHeight     =   1110
      ScaleWidth      =   2220
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   2280
   End
End
Attribute VB_Name = "frmItemSelectMulit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrKey As String
'��ڲ���
Private mstrTitle As String
Private mstrNote As String
Private mbytStyle As Byte
Private mstrSeek As String
Private mblnĩ�� As Boolean
Private mblnShowSub As Boolean
Private mblnShowRoot As Boolean
Private mblnMultiOne As Boolean

Private mstrSaveTag As String 'ע������ּ�
Private mstrSql As String
Private marrInput() As Variant

Private mblnSearch As Boolean '�Ƿ�ͨ�������кż���

Private mblnNoneWin As Boolean
Private mlngX As Long, mlngY As Long, mlngTxtH As Long
Private mblnMulitSelct As Boolean       '��ѡ

'���ڲ���
Private mrsSel As ADODB.Recordset
'�������
Private mblnOk As Boolean
 

Public Function ShowSelect(frmParent As Object, ByVal strSQL As String, bytStyle As Byte, _
    ByVal strTitle As String, blnĩ�� As Boolean, _
    ByVal strSeek As String, ByVal strNote As String, _
    ByVal blnShowSub As Boolean, blnShowRoot As Boolean, _
    ByVal blnNoneWin As Boolean, ByVal x As Long, _
    ByVal y As Long, ByVal txtH As Long, _
    ByRef Cancel As Boolean, ByVal blnMultiOne As Boolean, _
    ByVal blnSearch As Boolean, blnMulitSel As Boolean, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ��๦��ѡ����
'������
'     frmParent=��ʾ�ĸ�����
'     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
'     bytStyle=ѡ�������
'       Ϊ0ʱ:�б���:ID,��
'       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
'       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
'     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
'     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
'     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
'             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
'             bytStyle=1ʱ,�����Ǳ��������
'     strNote=ѡ������˵������
'     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
'     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
'     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
'     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
'     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
'     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
'     mblnMulitSelct-����ѡ�����(һ���ӵ���)
'     arrInput=��Ӧ�ĸ���SQL����ֵ,��˳����,����Ϊ��ȷ����
'���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
'˵����
'     1.ID���ϼ�ID����Ϊ�ַ�������
'     2.ĩ�����ֶβ�Ҫ����ֵ
'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    mblnMulitSelct = blnMulitSel
    mstrSql = strSQL
    If TypeName(arrInput) <> "Error" Then
        marrInput = arrInput
    Else
        marrInput = Array()
    End If
    
    mstrTitle = strTitle
    mstrNote = strNote
    mbytStyle = bytStyle
    mblnĩ�� = blnĩ��
    mstrSeek = strSeek
    mblnShowSub = blnShowSub
    mblnShowRoot = blnShowRoot
    mblnMultiOne = blnMultiOne
    mblnNoneWin = blnNoneWin
    mlngX = x: mlngY = y: mlngTxtH = txtH
    mblnSearch = blnSearch
    
    If Not frmParent Is Nothing Then
        mstrSaveTag = frmParent.Name & "_" & strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
    Else
        mstrSaveTag = strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
    End If
'    mblnMulitSelct = True
  
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOk Then
        Cancel = False
        Set ShowSelect = mrsSel
    Else
        Cancel = True
        Set ShowSelect = Nothing
    End If
End Function

Private Sub cmdCancel_Click()
    Set mrsSel = Nothing
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim lstItem As ListItem, objNode As Node
    Dim strIDIn As String
    strIDIn = ""
    If mblnMulitSelct Then
        If mbytStyle = 1 Then
            With tvw_s
                For Each objNode In tvw_s.Nodes
                    If objNode.Checked Then
                        strIDIn = strIDIn & "," & Mid(objNode.Key, 2)
                    End If
                Next
            End With
        Else
            With lvw
                For Each lstItem In lvw.ListItems
                    If lstItem.Checked Then
                        strIDIn = strIDIn & "," & Split(lstItem.Key, "_")(1)
                    End If
                Next
            End With
        End If
        If strIDIn <> "" Then
            strIDIn = Mid(strIDIn, 2)
        Else
            strIDIn = -1
        End If
        mrsSel.Filter = 0
        Set mrsSel = CopyNewRec(mrsSel, strIDIn)
        If mrsSel.RecordCount <> 0 Then mrsSel.MoveFirst
    Else
        If mrsSel.RecordCount = 0 Then Exit Sub
        If mblnĩ�� And mbytStyle = 1 Then
            If mrsSel!ĩ�� <> 1 Then Exit Sub
        End If
    End If
    mblnOk = True
    Unload Me
End Sub
Private Function CopyNewRec(ByVal rsSource As ADODB.Recordset, ByVal strIDIn As String) As ADODB.Recordset

    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer, lngLocate As Long
    
    lngLocate = -1
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        If .State = 1 Then .Close
        If rsSource.RecordCount <> 0 Then
            On Error Resume Next
            err = 0
            lngLocate = rsSource.AbsolutePosition
            If err <> 0 Then lngLocate = -1
            rsSource.MoveFirst
        End If
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Type = adNumeric Then
                .Fields.Append rsSource.Fields(intFields).Name, adDouble, rsSource.Fields(intFields).DefinedSize, adFldIsNullable        '0:��ʾ����
            Else
                .Fields.Append rsSource.Fields(intFields).Name, rsSource.Fields(intFields).Type, rsSource.Fields(intFields).DefinedSize, adFldIsNullable        '0:��ʾ����
            End If
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    If rsSource.RecordCount <> 0 Then rsSource.MoveFirst
    Do While Not rsSource.EOF
        If InStr(1, "," & strIDIn & ",", "," & zlStr.Nvl(rsSource!Id) & ",") > 0 Then
            rsTarget.AddNew
            For intFields = 0 To rsSource.Fields.Count - 1
                rsTarget.Fields(intFields) = rsSource.Fields(intFields).Value
            Next
            rsTarget.Update
        End If
        rsSource.MoveNext
    Loop
    
    If rsSource.RecordCount <> 0 Then rsSource.MoveFirst
    If lngLocate > 0 Then rsSource.Move lngLocate - 1
    Set CopyNewRec = rsTarget






'    Dim rsTarget As New ADODB.Recordset
'    Dim intFields As Integer, lngLocate As Long
'
'    lngLocate = -1
'    If rsSource.RecordCount <> 0 Then
'        On Error Resume Next
'        err = 0
'        lngLocate = rsSource.AbsolutePosition
'        If err <> 0 Then lngLocate = -1
'        rsSource.MoveFirst
'    End If
'
'    Set rsTarget = New ADODB.Recordset
'    With rsTarget
'        If .State = 1 Then .Close
'
'        err = 0: On Error GoTo 0
'        For intFields = 0 To rsSource.Fields.Count - 1
'            .Fields.Append rsSource.Fields(intFields).Name, rsSource.Fields(intFields).Type, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
'        Next
'
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockOptimistic
'        .Open
'
'        If rsSource.RecordCount <> 0 Then rsSource.MoveFirst
'        Do While Not rsSource.EOF
'            If InStr(1, "," & strIDIn & ",", "," & NVL(rsSource!Id) & ",") > 0 Then
'                .AddNew
'                For intFields = 0 To rsSource.Fields.Count - 1
'                        .Fields(intFields) = rsSource.Fields(intFields)
'                Next
'                .Update
'            End If
'            rsSource.MoveNext
'        Loop
'    End With
'    If rsSource.RecordCount <> 0 Then rsSource.MoveFirst
'    If lngLocate > 0 Then rsSource.Move lngLocate - 1
'    Set CopyNewRec = rsTarget
End Function

Private Sub Form_Activate()
    If lvw.Visible Then
        lvw.SetFocus
    Else
        tvw_s.SetFocus
    End If
    If mblnMulitSelct Then
        lblInfo.Caption = "��ѡ��һ����Ŀ������Ŀ,Ȼ����ȷ��"
    Else
        lblInfo.Caption = "��ѡ��һ����Ŀ,Ȼ����ȷ��"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And cmdOk.Enabled Then
        cmdOk_Click
    ElseIf KeyCode = vbKeyEscape And cmdCancel.Enabled Then
        cmdCancel_Click
    End If
End Sub

Private Function OpenSQLRecord(ByVal strSQL As String, strLog As String) As ADODB.Recordset
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim intMax As Integer, i As Integer
    Dim strSeq As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = marrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
    cmdData.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop

    '�����µĲ���
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = marrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = zlCommFun.ActualLen(varValue)
            If intMax = 0 Or intMax < 200 Then intMax = 200
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
    Next

    'ִ�з��ؼ�¼��
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, mstrTitle & "ѡ��", strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long
    Dim lngColW As Long, i As Integer
    Dim blnSame As Boolean, lngId As Long
    Dim strCode As String, strName As String
    Dim objNode As Node, strLog As String
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    If mblnMulitSelct = True Then
            lvw.Checkboxes = True
    End If
    
    mblnOk = False
    mstrKey = ""
        
    '��SQL���
    If UBound(marrInput) >= 0 Then
        Set mrsSel = OpenSQLRecord(mstrSql, strLog)
    Else
        Set mrsSel = New ADODB.Recordset
        mrsSel.CursorLocation = adUseClient
        Call SQLTest(App.ProductName, mstrTitle & "ѡ��", mstrSql)
        Set mrsSel = zlDatabase.OpenSQLRecord(mstrSql, "Form_Load")
        Call SQLTest
    End If
    
    'û�������򷵻�
    If mrsSel.EOF Then
        Screen.MousePointer = 0
        Set mrsSel = Nothing
        mblnOk = True: Unload Me: Exit Sub
    End If
     
    '����ƥ��ʱ�Զ����ص����
    If mstrSql Like "*%*" Or strLog Like "*%*" Then
        If mrsSel.RecordCount = 1 Then 'ֻ��һ������
            Screen.MousePointer = 0
            mblnOk = True: Unload Me: Exit Sub
        ElseIf mblnMultiOne And mbytStyle = 0 Then '������ͬ����
            blnSame = True
            For i = 1 To mrsSel.RecordCount
                If i = 1 Then
                    lngId = mrsSel!Id
                Else
                    If mrsSel!Id <> lngId Then blnSame = False: Exit For
                End If
                mrsSel.MoveNext
            Next
            mrsSel.MoveFirst
            If blnSame Then
                Screen.MousePointer = 0
                mblnOk = True: Unload Me: Exit Sub
            End If
        End If
    End If
    
    'ȷ�������ֶ�
    strCode = "": strName = ""
    For i = 0 To mrsSel.Fields.Count - 1
        If mrsSel.Fields(i).Name = "����" Then strCode = "����"
        If mrsSel.Fields(i).Name = "����" Then
            strName = mrsSel.Fields(i).Name
        ElseIf mrsSel.Fields(i).Name = "����" And strName = "" Then
            strName = mrsSel.Fields(i).Name
        End If
    Next
    If strName = "" Then strName = "����"
    
    '�������
    Select Case mbytStyle
        Case 0
            '������ͷ
            lvw.ColumnHeaders.Clear
            For i = 0 To mrsSel.Fields.Count - 1
                If (Not mrsSel.Fields(i).Name Like "*ID" Or mrsSel.Fields(i).Name = "����ID") And mrsSel.Fields(i).Name <> "ĩ��" Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.Fields(i).Name, mrsSel.Fields(i).Name
                    If mrsSel.Fields(i).Name Like "*��*" Or mrsSel.Fields(i).Name Like "*��*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Alignment = lvwColumnRight
                        
                    End If
                    
                End If
            Next
            If mblnSearch Then lvw.ColumnHeaders.Add , "_��", "��", , 2
            Call RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrSaveTag, "")
            lvw.Sorted = False
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Position = 1
            
            lvw.ListItems.Clear
            Call FillList
        Case 1
            '������������
            Set objNode = tvw_s.Nodes.Add(, , "Root", "����" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            objNode.Sorted = True
            
            If Not mrsSel.EOF Then
                For i = 1 To mrsSel.RecordCount
                    If strCode <> "" Then
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!Id, IIf(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.Fields(strName).Value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!Id, IIf(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.Fields(strName).Value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!Id, mrsSel.Fields(strName).Value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!Id, mrsSel.Fields(strName).Value, 1)
                        End If
                    End If
                    If objNode.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then
                        objNode.Selected = True
                        objNode.Parent.Expanded = True
                    End If
                    objNode.Sorted = True
                    mrsSel.MoveNext
                Next
                If tvw_s.SelectedItem.Index = 1 Then tvw_s.Nodes(1).Child.Selected = True
            End If
            tvw_s.SelectedItem.EnsureVisible
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        Case 2
            '��ĩ����������
            Set objNode = tvw_s.Nodes.Add(, , "Root", "����" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            objNode.Sorted = True
            
            If Not mrsSel.EOF Then
                mrsSel.Filter = "ĩ��=0"
                For i = 1 To mrsSel.RecordCount
                    If strCode <> "" Then
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!Id, IIf(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.Fields(strName).Value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!Id, IIf(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.Fields(strName).Value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!Id, mrsSel.Fields(strName).Value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!Id, mrsSel.Fields(strName).Value, 1)
                        End If
                    End If
                    objNode.Sorted = True
                    mrsSel.MoveNext
                Next
                If Not tvw_s.Nodes(1).Child Is Nothing Then tvw_s.Nodes(1).Child.Selected = True
            End If
            
            '������ͷ
            lvw.ColumnHeaders.Clear
            For i = 0 To mrsSel.Fields.Count - 1
                If (Not mrsSel.Fields(i).Name Like "*ID" Or mrsSel.Fields(i).Name = "����ID") And mrsSel.Fields(i).Name <> "ĩ��" Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.Fields(i).Name, mrsSel.Fields(i).Name
                    If mrsSel.Fields(i).Name Like "*��*" Or mrsSel.Fields(i).Name Like "*��*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Alignment = lvwColumnRight
                    End If
                End If
            Next
            If mblnSearch Then lvw.ColumnHeaders.Add , "_��", "��", , 2
            Call RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrSaveTag, "")
            lvw.Sorted = False
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Position = 1
            
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End Select
    
    '���ÿؼ��ɼ���
    '---------------------------------------------------------------
    If mstrTitle <> "" Then
        Me.Caption = mstrTitle & "ѡ��"
    End If
    If mstrNote <> "" Then
        lblInfo.Caption = mstrNote
    End If
    If mblnNoneWin Then
        pic.Width = 30
        pic.BackColor = vbBlack
        pic.ZOrder
        picInfo.Visible = False
        picCmd.Visible = False
        lvw.Appearance = ccFlat
        lvw.BorderStyle = ccFixedSingle
        tvw_s.Appearance = ccFlat
        tvw_s.BorderStyle = ccFixedSingle
    Else
        If mbytStyle <> 2 Then Me.Width = 4500 'ȱʡ���
        Call RestoreWinState(Me, App.ProductName, mstrSaveTag)
    End If
    Select Case mbytStyle
        Case 0
            lvw.Visible = True
            tvw_s.Visible = False
            pic.Visible = False
        Case 1
            lvw.Visible = False
            tvw_s.Visible = True
            pic.Visible = False
        Case 2
            lvw.Visible = True
            tvw_s.Visible = True
            pic.Visible = True
    End Select
    
    '��������ߴ�
    '---------------------------------------------------------------
    If mblnNoneWin Then
        Call FormSetCaption(Me, False, False)
        
        Me.Left = mlngX
        
        If mbytStyle = 1 Then
            Me.Width = 3100
        Else
            lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 75
            For i = 1 To lvw.ColumnHeaders.Count
                lngColW = lngColW + lvw.ColumnHeaders(i).Width
            Next
            If mbytStyle = 2 Then lngColW = lngColW + tvw_s.Width
            
            If Me.Left + lngColW + lngScrW > Screen.Width Then
                lngColW = 0
                For i = 1 To lvw.ColumnHeaders.Count
                    If lvw.ColumnHeaders(i).Width > 1800 Then lvw.ColumnHeaders(i).Width = 1800
                    lngColW = lngColW + lvw.ColumnHeaders(i).Width
                Next
                If Me.Left + lngColW + lngScrW > Screen.Width Then
                    Me.Width = Screen.Width - Me.Left
                Else
                    Me.Width = lngColW + lngScrW
                End If
            Else
                Me.Width = lngColW + lngScrW
            End If
        End If
        
        Me.Height = 3240
        lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '��Ļ���ø߶�
        If mlngY + mlngTxtH + Me.Height > lngScrH Then
            Me.Top = mlngY - Me.Height
        Else
            Me.Top = mlngY + mlngTxtH
        End If
        
        Call Form_Resize
    End If
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    Select Case mbytStyle
        Case 0 'ListView
            lvw.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            lvw.Left = 0
            lvw.Width = Me.ScaleWidth
            lvw.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
        Case 1
            tvw_s.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Left = 0
            tvw_s.Width = Me.ScaleWidth
            tvw_s.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
        Case 2
            tvw_s.Left = 0
            tvw_s.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
            
            pic.Top = tvw_s.Top
            pic.Height = tvw_s.Height
            lvw.Top = tvw_s.Top
            lvw.Height = tvw_s.Height
            
            If mblnNoneWin Then
                pic.Left = tvw_s.Width - pic.Width / 2
                lvw.Left = tvw_s.Width
                lvw.Width = Me.ScaleWidth - tvw_s.Width
            Else
                pic.Left = tvw_s.Width
                lvw.Left = tvw_s.Width + pic.Width
                lvw.Width = Me.ScaleWidth - tvw_s.Width - pic.Width
            End If
    End Select
    
    picBack.Left = lvw.Left
    picBack.Top = lvw.Top
    picBack.Width = lvw.Width
    picBack.Height = lvw.Height
    
    If Me.ScaleWidth - cmdCancel.Width * 1.3 >= cmdOk.Width + 700 Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.3
        cmdOk.Left = cmdCancel.Left - cmdOk.Width * 1.1
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub lvw_DblClick()
    If cmdOk.Enabled And Not lvw.SelectedItem Is Nothing Then cmdOk_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strFilter As String
    
    If mrsSel.Fields("ID").Type = adVarChar Then
        strFilter = "ID='" & Split(Item.Key, "_")(1) & "'"
    Else
        strFilter = "ID=" & Split(Item.Key, "_")(1)
    End If
    If mbytStyle = 2 Then strFilter = strFilter & " And ĩ��=1"
    
    mrsSel.Filter = strFilter
    cmdOk.Enabled = (mrsSel.RecordCount <> 0)
End Sub

Private Sub lvw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lvw_DblClick
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If mblnSearch Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If lvw.ListItems.Count >= CInt(strIdx) And CInt(strIdx) > 0 Then
                lvw.ListItems(CInt(strIdx)).Selected = True
                lvw.SelectedItem.EnsureVisible
                Call lvw_ItemClick(lvw.SelectedItem)
            End If
        End If
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvw_s.Width + x < 1000 Or lvw.Width - x < 1000 Then Exit Sub
        pic.Left = pic.Left + x
        tvw_s.Width = tvw_s.Width + x
        lvw.Left = lvw.Left + x
        lvw.Width = lvw.Width - x
        picBack.Left = picBack.Left + x
        picBack.Width = picBack.Width - x
        Me.Refresh
    End If
End Sub

Private Sub FillList()
'���ܣ�װ��ListView����
    Dim i As Integer, j As Integer
    Dim objItem As ListItem
        
    lvw.Visible = False
    Screen.MousePointer = 11
    For i = 1 To mrsSel.RecordCount
        For j = 0 To mrsSel.Fields.Count - 1
            If (Not mrsSel.Fields(j).Name Like "*ID" Or mrsSel.Fields(j).Name = "����ID") And mrsSel.Fields(j).Name <> "ĩ��" Then
                If lvw.ColumnHeaders("_" & mrsSel.Fields(j).Name).Index = 1 Then
                    If mblnSearch Then '�ؼ��ּ����к�
                        Set objItem = lvw.ListItems.Add(, i & "_" & mrsSel!Id, IIf(IsNull(mrsSel.Fields(j).Value), "", mrsSel.Fields(j).Value), , 1)
                    Else
                        Set objItem = lvw.ListItems.Add(, "_" & mrsSel!Id, IIf(IsNull(mrsSel.Fields(j).Value), "", mrsSel.Fields(j).Value), , 1)
                    End If
                    If objItem.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then objItem.Selected = True
                Else
                    objItem.SubItems(lvw.ColumnHeaders("_" & mrsSel.Fields(j).Name).Index - 1) = IIf(IsNull(mrsSel.Fields(j).Value), "", mrsSel.Fields(j).Value)
                End If
                objItem.Checked = False
                
            End If
        Next
        If mblnSearch Then objItem.SubItems(lvw.ColumnHeaders("_��").Index - 1) = i
        mrsSel.MoveNext
    Next
    
    Call zlControl.LvwSetColWidth(lvw)
    '20031013:���������
    If lvw.Width > Screen.Width / 2 Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width > 1800 Then lvw.ColumnHeaders(i).Width = 1800
        Next
    End If
    
    If Not lvw.SelectedItem Is Nothing Then
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
    lvw.Refresh
    lvw.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub tvw_s_DblClick()
    If cmdOk.Enabled And mbytStyle = 1 Then cmdOk_Click
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim mstrKeys As String, i As Integer
    Dim strFilter As String
    
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    If mbytStyle = 1 Then
        If Node.Key <> "Root" Then
            If mrsSel.Fields("ID").Type = adVarChar Then
                mrsSel.Filter = "ID='" & Mid(Node.Key, 2) & "'"
            Else
                mrsSel.Filter = "ID=" & Mid(Node.Key, 2)
            End If
            If mblnĩ�� Then
                cmdOk.Enabled = (mrsSel!ĩ�� = 1)
            Else
                cmdOk.Enabled = True
            End If
        Else
            cmdOk.Enabled = False
        End If
    ElseIf mbytStyle = 2 Then
        lvw.ListItems.Clear
        If Node.Key = "Root" Then
            If mblnShowRoot Then
                mrsSel.Filter = "ĩ��=1" '��������ʱ����
            Else
                mrsSel.Filter = "ĩ��=-1"
            End If
            If Visible Then lvw.SetFocus
        Else
            If mblnShowSub Then
                mstrKeys = GetSubTree(Node) '��������ʱ����
            Else
                mstrKeys = Mid(Node.Key, 2)
            End If
            For i = 0 To UBound(Split(mstrKeys, ","))
                If mrsSel.Fields("�ϼ�ID").Type = adVarChar Then
                    strFilter = strFilter & " Or (ĩ��=1 And �ϼ�ID='" & Split(mstrKeys, ",")(i) & "')"
                Else
                    strFilter = strFilter & " Or (ĩ��=1 And �ϼ�ID=" & Split(mstrKeys, ",")(i) & ")"
                End If
            Next
            strFilter = Mid(strFilter, 5)
            mrsSel.Filter = strFilter
            
'            If mrsSel.Fields("�ϼ�ID").Type = adVarChar Then
'                mrsSel.Filter = "ĩ��=1 And �ϼ�ID='" & Mid(Node.Key, 2) & "'"
'            Else
'                mrsSel.Filter = "ĩ��=1 And �ϼ�ID=" & Mid(Node.Key, 2)
'            End If
        End If
        If Not mrsSel.EOF Then Call FillList
    End If
End Sub

Private Function GetSubTree(ByVal objNode As Node) As String
'���ܣ�����һ��������������Key(���ý��)
    Dim mstrKeys As String
    Dim objTmp As Node
    
    mstrKeys = "," & Mid(objNode.Key, 2) & mstrKeys
    Set objTmp = objNode.Child
    Do While Not objTmp Is Nothing
        If objTmp.Children > 0 Then
            mstrKeys = "," & GetSubTree(objTmp) & mstrKeys
        Else
            mstrKeys = "," & Mid(objTmp.Key, 2) & mstrKeys
        End If
        Set objTmp = objTmp.Next
    Loop
    GetSubTree = Mid(mstrKeys, 2)
End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If mblnSearch And ColumnHeader.Key = "_��" Then Exit Sub
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
        
    If mblnSearch Then
        For intIdx = 1 To lvw.ListItems.Count
            lvw.ListItems(intIdx).SubItems(lvw.ColumnHeaders("_��").Index - 1) = intIdx
        Next
    End If
    intIdx = ColumnHeader.Index
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub


