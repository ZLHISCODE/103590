VERSION 5.00
Begin VB.Form frmDictEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmDictEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cboEdit 
      Height          =   300
      Index           =   0
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmd�ϼ� 
      Caption         =   "��"
      Height          =   270
      Left            =   2430
      TabIndex        =   8
      Top             =   1875
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CheckBox Chk�Ƿ� 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   7
      Top             =   2445
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CheckBox chkĩ�� 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   285
      TabIndex        =   6
      Top             =   3105
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Frame fraSplit 
      Height          =   4485
      Left            =   2700
      TabIndex        =   5
      Top             =   -510
      Width           =   30
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   420
      TabIndex        =   4
      Top             =   1860
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2970
      TabIndex        =   2
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2970
      TabIndex        =   1
      Top             =   180
      Width           =   1100
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "Check1"
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblCombox 
      AutoSize        =   -1  'True
      Caption         =   "Combox"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmDictEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrOwner As String       '��ǰ�༭�������������
Private mstrTable As String       '��ǰ�༭�ı���
Private mstr���� As String        '��ǰ�༭�ļ�¼��ʶ
Private mint����  As Integer      '�����ֶε����
Private mint����  As Integer      '�����ֶε����
Private mint����  As Integer      '�����ֶε����
Private mint���볤��  As Integer  '���õ�Դ
Private mstr�ϼ� As String        '���ӡ��޸Ľ���ʱ�������ϼ�ID 2010-04-06
Private mvar���ӹ�ϵ As Variant

Private mlng����() As Long        '�ֶ�����,Ϊ1��ʾ������,2��ʾ����
Private mblnChange As Boolean
Private mblnRISChange As Boolean

Private Sub cboEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub cmd�ϼ�_Click()
    Dim vRect As Rect
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim rtn As String, CCtype As ChooseColorType, i As Integer
    
    On Error GoTo ErrH
    If cmd�ϼ�.Tag = "��ɫ" Then
        For i = txtEdit.LBound To txtEdit.UBound
            If txtEdit(i).Tag = "��ɫ" Then Exit For
        Next
        With CCtype
            .lStructSize = Len(CCtype)
            .hwndOwner = Me.hWnd
            .hInstance = App.hInstance
            .flags = 0
            .lpCustColors = String$(16 * 16, 0)
        End With
        rtn = ChooseColor(CCtype)
        If rtn >= 1 Then
            txtEdit(i).Text = CCtype.rgbResult
            txtEdit(i).ForeColor = CCtype.rgbResult
        Else
            txtEdit(i).Text = 0
            txtEdit(i).ForeColor = 0
        End If
    Else
        vRect = zlControl.GetControlRect(txtEdit(cmd�ϼ�.Tag).hWnd)
        
        gstrSQL = "Select * From (select '0' as ID,null as �ϼ�ID,'' as ����,'ȫ��' as ����,0 as ĩ�� From dual " & _
                  "union all Select to_char(����) as ID,nvl(�ϼ�,0) As �ϼ�ID, to_char(����) as ����, ����, ĩ�� " & _
                  " From " & mstrOwner & "." & mstrTable & " Where nvl(ĩ��,0)=0 ) Order by nvl(�ϼ�ID,0),Id "
        '��ʾѡ����
        Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 1, "��Ŀ", , , , , , False, vRect.Left, vRect.Top, txtEdit(cmd�ϼ�.Tag).Height, blnCancel, , True)
                
        If Not blnCancel Then
            If Not rsTmp Is Nothing Then
                txtEdit(cmd�ϼ�.Tag).Tag = IIf(txtEdit(cmd�ϼ�.Tag).Text = "", "ȫ��", txtEdit(cmd�ϼ�.Tag).Text)
                txtEdit(cmd�ϼ�.Tag).Text = IIf(IsNull(rsTmp("����")), "", rsTmp("����"))
                'ͬʱ�ı�mstr�ϼ���ֵ 2010-04-06
                mstr�ϼ� = IIf(IsNull(rsTmp("����")), "", rsTmp("����"))
            End If
        End If
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    txtEdit(mint����).SetFocus
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
    If Save����() = False Then Exit Sub
    Call frmDictManager.FillList
    If mstr���� <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstr���� = ""
    chkLog.Value = 0
    For i = 1 To lblEdit.Count - 1
        '���ϼ��⣬����ȫ����� 2010-04-06
        If Left(lblEdit(i).Caption, 2) = "�ϼ�" Then
            txtEdit(i).Text = mstr�ϼ�
        Else
            txtEdit(i).Text = ""
        End If
    Next
    If mstr���� = "" Then txtEdit(mint����).Text = zlDatabase.GetMax(mstrOwner & "." & mstrTable, "����", mint���볤��)
    mblnChange = False
    txtEdit(mint����).SetFocus
'    Unload Me
End Sub

Private Function IsValid() As Boolean
'����:����������������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = 1 To lblEdit.Count - 1
        strTemp = Trim(txtEdit(i).Text)
        If zlCommFun.StrIsValid(strTemp, txtEdit(i).MaxLength, txtEdit(i).hWnd) = False Then
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
        If InStr(txtEdit(i).Text, ",") > 0 Or InStr(txtEdit(i).Text, ";") > 0 Then
            MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            Exit Function
        End If
        If i = mint���� Or i = mint���� Then
            If Len(strTemp) = 0 Then
                MsgBox lblEdit(i).Tag & "����Ϊ�ա�", vbExclamation, gstrSysName
                txtEdit(i).Text = ""
                txtEdit(i).SetFocus
                Exit Function
            End If
        Else
            '�жϸ��ֵ������Ƿ���Check is not nullԼ����Ŀǰֻ��Edit�ؼ��жϡ�
            If IsCheckConstraint(mstrOwner, mstrTable, lblEdit(i).Tag, 2) And Trim(txtEdit(i).Text) = "" Then
                MsgBox lblEdit(i).Tag & "����Ϊ�ա�", vbExclamation, gstrSysName
                txtEdit(i).SetFocus
                Exit Function
            End If
        End If
        If mlng����(i) = 1 Then
            '�������ֶ�
            If strTemp <> "" And Not IsNumeric(strTemp) Then
                MsgBox lblEdit(i).Tag & "Ӧ���������֡�", vbExclamation, gstrSysName
                zlControl.TxtSelAll txtEdit(i)
                txtEdit(i).SetFocus
                Exit Function
            End If
        
        End If
        If mlng����(i) = 2 Then
            '�������ֶ�
            strTemp = zlCommFun.AddDate(strTemp)
            
            If strTemp <> "" Then
                If Not IsDate(strTemp) Then
                    MsgBox lblEdit(i).Tag & "�������ڸ�ʽ(yyyy-mm-dd)��(yyyymmdd)��", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    txtEdit(i).SetFocus
                    Exit Function
                End If
                If zlCommFun.ActualLen(strTemp) <> 10 Then
                    MsgBox lblEdit(i).Tag & "���Ȳ���,Ӧ��Ϊ10λ(yyyy-mm-dd)��", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    txtEdit(i).SetFocus
                    Exit Function
                End If
                Err = 0
                On Error Resume Next
                strTemp = Format(strTemp, "yyyy-mm-dd")
                If Err <> 0 Or Not IsDate(strTemp) Then
                    MsgBox lblEdit(i).Tag & "�������ڸ�ʽ(yyyy-mm-dd)��", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    txtEdit(i).SetFocus
                    Exit Function
                End If
                
                txtEdit(i).Text = strTemp
            End If
        End If
    Next
    
    If chkĩ��.Visible = True Then
        If chkĩ��.Value <> 1 And chkLog.Value = 1 Then
            MsgBox "ֻ��ĩ����Ŀ��������Ϊȱʡֵ��", vbInformation, gstrSysName
            chkLog.Value = 0
            Exit Function
        End If
    End If

    IsValid = True
End Function

Private Function Save����() As Boolean
'����:����������ݽ��б���
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim strSQL As String
    Dim strTemp As String
    Dim i As Long
    Dim lngSystem As Long
    Dim str���� As String, str���� As String
    Dim str�ϼ� As String
    Dim blnTrans As Boolean, lngReturn As Long
    
    With frmDictManager.cmbSys
        lngSystem = .ItemData(.ListIndex) \ 100
    End With
    
    On Error GoTo errHandle
    If mstr���� = "" Then       '����һ����¼
        strSQL = "insert into " & mstrOwner & "." & mstrTable & " ("
        For i = 1 To lblEdit.Count - 1
            strSQL = strSQL & lblEdit(i).Tag & ","
            If mlng����(i) = 2 Then
                strTemp = strTemp & "to_Date('" & Format(Trim(txtEdit(i).Text), "yyyy-mm-dd") & "','yyyy-mm-dd'),"
            Else
                strTemp = strTemp & "'" & Trim(txtEdit(i).Text) & "',"
                If mstrTable = "���쳣��ԭ��" And lblEdit(i).Tag = "����" Then str���� = Trim(txtEdit(i).Text)
                If mstrTable = "���쳣��ԭ��" And lblEdit(i).Tag = "�ϼ�" Then str�ϼ� = Trim(txtEdit(i).Text)
            End If
        Next
        
        For i = 1 To Chk�Ƿ�.Count - 1
            strSQL = strSQL & Chk�Ƿ�(i).Tag & ","
            strTemp = strTemp & IIf(Chk�Ƿ�(i).Value = 1, "1,", "0,")
        Next
        
        For i = 1 To cboEdit.Count - 1
            If mvar���ӹ�ϵ(1) = "����" Then
                strSQL = strSQL & lblCombox(i).Tag & ","
                strTemp = strTemp & "'" & Mid(cboEdit(i).Text, InStr(cboEdit(i).Text, "-") + 1, Len(cboEdit(i).Text)) & "',"
            ElseIf mvar���ӹ�ϵ(1) = "����" Then
                strSQL = strSQL & lblCombox(i).Tag & ","
                If InStr(cboEdit(i).Text, "-") > 0 Then
                    strTemp = strTemp & "'" & Mid(cboEdit(i).Text, 1, InStr(cboEdit(i).Text, "-") - 1) & "',"
                Else
                    strTemp = strTemp & "'',"
                End If
            Else
                strSQL = strSQL & lblCombox(i).Tag & ","
                strTemp = strTemp & cboEdit(i).ItemData(cboEdit(i).ListIndex) & ","
            End If
        Next
        
        If chkĩ��.Tag <> "" Then
            strSQL = strSQL & chkĩ��.Tag & ","
            strTemp = strTemp & IIf(chkĩ��.Value = 1, "1,", "0,")
        End If
        
        If chkLog.Visible = False Then
            strSQL = Left(strSQL, Len(strSQL) - 1)
            strTemp = Left(strTemp, Len(strTemp) - 1)
        Else
            strSQL = strSQL & chkLog.Tag
            strTemp = strTemp & IIf(chkLog.Value = 1, "1", "0")
        End If
        
        If mstrTable = "���쳣��ԭ��" And InStr(strSQL, "����") = 0 Then
            strSQL = strSQL & ",����"
            strTemp = strTemp & ",(Select ���� From " & mstrOwner & "." & mstrTable & " Where ����='" & str�ϼ� & "')"
        End If
        
        strSQL = strSQL & ") values ( " & strTemp & ")"
    Else    '�޸�
        strSQL = "update " & mstrOwner & "." & mstrTable & " set "
        For i = 1 To lblEdit.Count - 1
            If mlng����(i) = 2 Then
                strSQL = strSQL & lblEdit(i).Tag & "=" & "to_Date('" & Format(Trim(txtEdit(i).Text), "yyyy-mm-dd") & "','yyyy-mm-dd'),"
            Else
                strSQL = strSQL & lblEdit(i).Tag & "=" & "'" & Trim(txtEdit(i).Text) & "',"
                If mstrTable = "���쳣��ԭ��" And lblEdit(i).Tag = "����" Then str���� = Trim(txtEdit(i).Text)
                If mstrTable = "���쳣��ԭ��" And lblEdit(i).Tag = "�ϼ�" Then str�ϼ� = Trim(txtEdit(i).Text)
            End If
        Next
        
        For i = 1 To Chk�Ƿ�.Count - 1
            strSQL = strSQL & Chk�Ƿ�(i).Tag & "=" & IIf(Chk�Ƿ�(i).Value = 1, "1,", "0,")
        Next
        
        For i = 1 To cboEdit.Count - 1
            If mvar���ӹ�ϵ(1) = "����" Then
                strSQL = strSQL & lblCombox(i).Tag & "='" & Mid(cboEdit(i).Text, InStr(cboEdit(i).Text, "-") + 1, Len(cboEdit(i).Text)) & "',"
            ElseIf mvar���ӹ�ϵ(1) = "����" Then
                strSQL = strSQL & lblCombox(i).Tag & "='" & Mid(cboEdit(i).Text, 1, InStr(cboEdit(i).Text, "-") - 1) & "',"
            Else
                strSQL = strSQL & lblCombox(i).Tag & "=" & cboEdit(i).ItemData(cboEdit(i).ListIndex) & ","
                If mstrTable = "���쳣��ԭ��" And lblCombox(i).Tag = "����" Then str���� = cboEdit(i).ItemData(cboEdit(i).ListIndex)
            End If
        Next
        
        If chkĩ��.Tag <> "" Then
            strSQL = strSQL & chkĩ��.Tag & "=" & IIf(chkĩ��.Value = 1, "1,", "0,")
        End If
        
        If chkLog.Visible = False Then
            strSQL = Left(strSQL, Len(strSQL) - 1)
        Else
            strSQL = strSQL & chkLog.Tag & "=" & IIf(chkLog.Value = 1, "1", "0")
        End If
        If mstrTable = "���쳣��ԭ��" And InStr(strSQL, "����") = 0 Then
            strSQL = strSQL & " ,���� = (Select ���� From " & mstrOwner & "." & mstrTable & " Where ���� = '" & str�ϼ� & "' And Rownum = 1) "
        End If
        strSQL = strSQL & " where ���� = '" & mstr���� & "'"
    
    End If
    gcnOracle.BeginTrans: blnTrans = True
    If chkLog.Tag = "ȱʡ��־" And chkLog.Value = 1 Then
        strTemp = "update " & mstrOwner & "." & mstrTable & " set ȱʡ��־=0"
        '�ù��̽��з�װ
        gstrSQL = "ZL_�ֵ����_execute('" & Replace(strTemp, "'", "''") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    '�ù��̽��з�װ
    gstrSQL = "ZL_�ֵ����_execute('" & Replace(strSQL, "'", "''") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If frmDictManager.gblnHaveRIS And mblnRISChange Then
        If mstr���� <> txtEdit(mint����).Text And mstr���� <> "" Then '����仯������ɾ��
            lngReturn = frmDictManager.gobjRIS.HISBasicDictTable(Decode(mstrTable, "�ѱ�", 4, "ҽ�Ƹ��ʽ", 5, "����", 6, "����״��", 7, "ְҵ", 8, "�Ա�", 9), 3, mstr����)
            lngReturn = frmDictManager.gobjRIS.HISBasicDictTable(Decode(mstrTable, "�ѱ�", 4, "ҽ�Ƹ��ʽ", 5, "����", 6, "����״��", 7, "ְҵ", 8, "�Ա�", 9), 1, txtEdit(mint����).Text)
        ElseIf mstr���� <> "" Then
            lngReturn = frmDictManager.gobjRIS.HISBasicDictTable(Decode(mstrTable, "�ѱ�", 4, "ҽ�Ƹ��ʽ", 5, "����", 6, "����״��", 7, "ְҵ", 8, "�Ա�", 9), 2, mstr����)
        Else
            lngReturn = frmDictManager.gobjRIS.HISBasicDictTable(Decode(mstrTable, "�ѱ�", 4, "ҽ�Ƹ��ʽ", 5, "����", 6, "����״��", 7, "ְҵ", 8, "�Ա�", 9), 1, txtEdit(mint����).Text)
        End If
        If lngReturn <> 1 And frmDictManager.gblnMustRIS Then
            gcnOracle.RollbackTrans
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISBasicDictTable)δ���óɹ������ܽ��е�ǰ������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If str���� <> "" Then
        strTemp = "update " & mstrOwner & "." & mstrTable & " set ����= " & str���� & " Where �ϼ� = '" & str���� & "'"
        '�ù��̽��з�װ
        gstrSQL = "ZL_�ֵ����_execute('" & Replace(strTemp, "'", "''") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    gcnOracle.CommitTrans: blnTrans = False
    If chkĩ��.Tag <> "" Then
        If txtEdit(cmd�ϼ�.Tag).Tag <> "" Then
            '�����ϼ�
            Call UpdateMain(0)
        Else
            Call UpdateMain(IIf(chkĩ��.Value = 1, "1", "0"))
        End If
    Else
        Call UpdateMain(1)
    End If
    Save���� = True
    Exit Function

errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub UpdateMain(ByVal strĩ�� As String)
'���ܣ�����������
    Dim lst As ListItem
    Dim ch As ColumnHeader
    Dim lngCount As Long
    Dim strTemp As String
    Dim intNodesOld As Integer
    
    If strĩ�� = 0 Then
        intNodesOld = frmDictManager.tvwMain.SelectedItem.Index
        Call frmDictManager.frmRefresh
        frmDictManager.TreeViewExpand frmDictManager.tvwMain, True
        frmDictManager.tvwMain.Nodes(intNodesOld).Selected = True
        Exit Sub
    End If
    
    With frmDictManager.lvwMain
        If mstr���� = "" Then
'            If strĩ�� = 1 Then
                Set lst = .ListItems.Add(, "C" & txtEdit(mint����).Text, txtEdit(mint����).Text, "Item", "Item")
                If .ListItems.Count = 1 Then
                    lst.Selected = True
                End If
'            Else
'                '����һ�����
'            End If
        Else
            If mstr���� <> txtEdit(mint����).Text Then
                
                '����ı䣬��Ҫ�޸���Keyֵ
                .ListItems.Remove .SelectedItem.Key
                Set lst = .ListItems.Add(, "C" & txtEdit(mint����).Text, txtEdit(mint����).Text, "Item", "Item")
                lst.Selected = True
                lst.EnsureVisible
            Else
                Set lst = .SelectedItem
                lst.Text = txtEdit(mint����).Text
            End If
        End If
        
        For Each ch In .ColumnHeaders
            strTemp = ch.Text
            If strTemp <> "����" Then
                For lngCount = 1 To lblEdit.Count - 1
                    If strTemp = lblEdit(lngCount).Tag Then '��ʾ��ͬ�ֶ�
                        Exit For
                    End If
                Next
                
                If lngCount < lblEdit.Count Then
                    '�ڱ༭�����ҵ�
                    If mlng����(lngCount) = 2 Then
                        lst.SubItems(ch.SubItemIndex) = Format(Trim(txtEdit(lngCount).Text), "yyyy-mm-dd")
                    Else
                        If lblEdit(lngCount).Tag = "�ϼ�" Then
                            lst.SubItems(ch.SubItemIndex) = txtEdit(lngCount).Tag
                        Else
                            lst.SubItems(ch.SubItemIndex) = txtEdit(lngCount).Text
                        End If
                    End If
                Else
                    If strTemp = "ȱʡ��־" Then
                        If chkLog.Value = 1 Then
                            '��ListView�и��е�ֵȫ���
                            For lngCount = 1 To .ListItems.Count
                                .ListItems(lngCount).SubItems(ch.SubItemIndex) = ""
                            Next
                        End If
                        lst.SubItems(ch.SubItemIndex) = IIf(chkLog.Value = 1, "��", "")
                    End If
 
                End If
                Dim intChk As Integer
                If strTemp Like "�Ƿ�*" Then
                    For intChk = 1 To Chk�Ƿ�.Count - 1
                        If strTemp = Chk�Ƿ�(intChk).Tag Then
                            lst.SubItems(ch.SubItemIndex) = IIf(Chk�Ƿ�(intChk).Value = 1, "��", "")
                        End If
                    Next
                End If
            End If
        Next
    End With
    Call frmDictManager.SetMenu
End Sub

Public Function �༭����(ByVal strOwner As String, ByVal strTable As String, Optional str���� As String = "", Optional intĩ�� As Integer = -1, Optional str�ϼ� As String) As Boolean
'����:��������ô��ڽ���ͨѶ�ĳ���
'����:strTable  Ҫ�༭�ı���
'     str����     Ҫ�༭�ı�����ؼ���
'����ֵ:�ɹ�����True,����ΪFalse
    Dim rs����� As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim fld As Field
    Dim lst As ListItem
    Dim sngY As Single     '��ǰ�༭��ĸ߶�
    Dim sngMaxW As Single  '�༭��������
    Dim intTemp As Integer, intChkTmp As Integer, lngcboTmp As Long
    Dim strTmp As String
    Dim i As Long
    Dim blnRISChange As Boolean, blnTrans As Boolean
    '��ʼ������
    sngY = 200
    sngMaxW = 0
    mstrOwner = strOwner
    mstrTable = strTable
    mstr���� = str����
    mblnRISChange = False
    If mstrOwner = frmDictManager.gstrSTOwner Then
        '֪ͨRIS������䶯
        '�ѱ����ʱû��ͨ���ù��߹����Ա������״��Ϊ�̶���
        If InStr(",�ѱ�,ҽ�Ƹ��ʽ,����,����״��,ְҵ,�Ա�,", "," & strTable & ",") > 0 Then
            mblnRISChange = True
            If frmDictManager.gblnMustRIS And Not frmDictManager.gblnHaveRIS Then
                MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ������ֵ��" & strTable & "���е����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    '�����ϼ��ִ� 2010-04-06
    If str�ϼ� <> "oot" Then
        mstr�ϼ� = str�ϼ�
    Else
        mstr�ϼ� = ""
    End If
    
    mint���볤�� = 0
    mint���� = 0
    mint���� = 0
    chkLog.Tag = ""
    chkĩ��.Tag = ""
    
    '���ӱ������Ӧ�ֶ�����
    strTmp = IsPathProperty(strOwner, strTable)
    If strTable = "ҽ�ƻ���" Or strTable = "��Ժת��" Then strTmp = ";"
    mvar���ӹ�ϵ = Split(strTmp, ";")
    If UBound(mvar���ӹ�ϵ) >= 2 Then
        If mvar���ӹ�ϵ(2) = "����" Then '�����ϼ����벻�������б�չʾ�����ܻ�����������Լ��༭������
            mvar���ӹ�ϵ = Split(";", ";")
        End If
    End If
    On Error Resume Next
    rs�����.CursorLocation = adUseClient
    
    gstrSQL = "select * from " & strOwner & "." & strTable & " where ���� = [1]"
    Set rs����� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����)
    If Err.Number <> 0 Then Err.Clear '���ܲ�ѯ��������
    ReDim mlng����(0 To rs�����.Fields.Count)
    For Each fld In rs�����.Fields
        If UCase(fld.Name) = "��ԴID" And UBound(mvar���ӹ�ϵ) >= 2 Then
            If UCase(mvar���ӹ�ϵ(2)) = "RESOURCEINFO" Then GoTo makContinue
        ElseIf fld.Name = "ȱʡ��־" Then
            '���߼�����
            chkLog.Caption = fld.Name
            chkLog.Tag = fld.Name
            chkLog.Caption = fld.Name & IIf(fld.Name = "ȱʡ��־", "��ע�⣺�����־���������ԣ�", "")
            chkLog.Left = 200
            chkLog.Width = 300 + Me.TextWidth(chkLog.Caption)
            If chkLog.Width + 200 > sngMaxW Then sngMaxW = chkLog.Width + 200
            chkLog.Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)
            chkLog.Visible = True
            
        ElseIf fld.Name Like "�Ƿ�*" Then
            intChkTmp = Chk�Ƿ�.Count
            Load Chk�Ƿ�(intChkTmp)
            Chk�Ƿ�(intChkTmp).Caption = fld.Name
            Chk�Ƿ�(intChkTmp).Tag = fld.Name
            Chk�Ƿ�(intChkTmp).Left = 200
            Chk�Ƿ�(intChkTmp).Width = 300 + Me.TextWidth(Chk�Ƿ�(intChkTmp).Caption)
            If Chk�Ƿ�(intChkTmp).Width + 200 > sngMaxW Then sngMaxW = Chk�Ƿ�(intChkTmp).Width + 200
            Chk�Ƿ�(intChkTmp).Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)

            Chk�Ƿ�(intChkTmp).Top = sngY
            sngY = sngY + Chk�Ƿ�(intChkTmp).Height + 100
            If Chk�Ƿ�(intChkTmp).Width + Chk�Ƿ�(intChkTmp).Left > sngMaxW Then sngMaxW = Chk�Ƿ�(intChkTmp).Width + Chk�Ƿ�(intChkTmp).Left

            Chk�Ƿ�(intChkTmp).Visible = True
            
        ElseIf fld.Name = "ĩ��" Then
            chkĩ��.Caption = fld.Name
            chkĩ��.Tag = fld.Name
            chkĩ��.Left = 200
            chkĩ��.Width = 300 + Me.TextWidth(chkĩ��.Caption)
            If chkĩ��.Width + 200 > sngMaxW Then sngMaxW = chkĩ��.Width + 200
            If intĩ�� <> -1 Then
                chkĩ��.Value = IIf(IIf(IsNull(intĩ��), 0, intĩ��), 1, 0)
            Else
                chkĩ��.Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)
            End If
        
        ElseIf mvar���ӹ�ϵ(0) = fld.Name Then
            lngcboTmp = lblCombox.Count
            Load lblCombox(lngcboTmp)
            Load cboEdit(lngcboTmp)
            lblCombox(lngcboTmp).Top = sngY
            lblCombox(lngcboTmp).Left = 200
            lblCombox(lngcboTmp).Tag = fld.Name
            lblCombox(lngcboTmp).Caption = fld.Name & "(&" & intTemp + lngcboTmp & ")"
            cboEdit(lngcboTmp).Top = lblCombox(lngcboTmp).Top
            cboEdit(lngcboTmp).Left = lblCombox(lngcboTmp).Left + lblCombox(lngcboTmp).Width + 100
            '����cboEdit(0)������
            'Call SetPathProp(lngcboTmp)
            Call SetSelectProp(lngcboTmp)
            
            If mvar���ӹ�ϵ(1) = "����" Then
                strTmp = "select ���� || '-' || ���� ���� from " & mvar���ӹ�ϵ(2) & " where " & mvar���ӹ�ϵ(1) & "=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "��ȡ��������", fld.Value)
                cboEdit(lngcboTmp).Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
            Else
                For i = 0 To cboEdit(lngcboTmp).ListCount - 1
                    If cboEdit(lngcboTmp).ItemData(i) = fld.Value Then
                        cboEdit(lngcboTmp).Text = cboEdit(lngcboTmp).List(i)
                        Exit For
                    End If
                Next
            End If
            sngY = sngY + cboEdit(lngcboTmp).Height + 100
            
            cboEdit(lngcboTmp).Visible = True
            lblCombox(lngcboTmp).Visible = True
        ElseIf fld.Name = "����" And strTable = "���쳣��ԭ��" Then
            If intĩ�� = 0 Then
                lngcboTmp = lblCombox.Count
                Load lblCombox(lngcboTmp)
                Load cboEdit(lngcboTmp)
                lblCombox(lngcboTmp).Top = sngY
                lblCombox(lngcboTmp).Left = 200
                lblCombox(lngcboTmp).Tag = fld.Name
                lblCombox(lngcboTmp).Caption = fld.Name & "(&" & intTemp + lngcboTmp & ")"
                cboEdit(lngcboTmp).Top = lblCombox(lngcboTmp).Top
                cboEdit(lngcboTmp).Left = lblCombox(lngcboTmp).Left + lblCombox(lngcboTmp).Width + 100
                
                '0-δ�����ԭ��;1-���������ԭ��;2-�����˳���ԭ��
                cboEdit(lngcboTmp).Clear
                cboEdit(lngcboTmp).AddItem "0-δ�����ԭ��": cboEdit(lngcboTmp).ItemData(0) = 0
                cboEdit(lngcboTmp).AddItem "1-���������ԭ��": cboEdit(lngcboTmp).ItemData(1) = 1
                cboEdit(lngcboTmp).AddItem "2-�����˳���ԭ��": cboEdit(lngcboTmp).ItemData(2) = 2
                
                For i = 0 To cboEdit(lngcboTmp).ListCount - 1
                    If cboEdit(lngcboTmp).ItemData(i) = fld.Value Then
                        cboEdit(lngcboTmp).Text = cboEdit(lngcboTmp).List(i)
                        Exit For
                    End If
                Next
                sngY = sngY + cboEdit(lngcboTmp).Height + 100
                
                cboEdit(lngcboTmp).Visible = True
                lblCombox(lngcboTmp).Visible = True
            End If
        ElseIf fld.Type = adNumeric And fld.Precision = 1 Then
            'Numeric���ͣ����1B����CheckԼ������CheckBox���֡�����������д����Ҫ����ִ��Ч�ʡ�
            If IsCheckConstraint(mstrOwner, strTable, fld.Name, 1) = True Then
                intChkTmp = Chk�Ƿ�.Count
                Load Chk�Ƿ�(intChkTmp)
                Chk�Ƿ�(intChkTmp).Caption = fld.Name
                Chk�Ƿ�(intChkTmp).Tag = fld.Name
                Chk�Ƿ�(intChkTmp).Left = 200
                Chk�Ƿ�(intChkTmp).Width = 300 + Me.TextWidth(Chk�Ƿ�(intChkTmp).Caption)
                If Chk�Ƿ�(intChkTmp).Width + 200 > sngMaxW Then sngMaxW = Chk�Ƿ�(intChkTmp).Width + 200
                Chk�Ƿ�(intChkTmp).Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)
    
                Chk�Ƿ�(intChkTmp).Top = sngY
                sngY = sngY + Chk�Ƿ�(intChkTmp).Height + 100
                If Chk�Ƿ�(intChkTmp).Width + Chk�Ƿ�(intChkTmp).Left > sngMaxW Then sngMaxW = Chk�Ƿ�(intChkTmp).Width + Chk�Ƿ�(intChkTmp).Left
    
                Chk�Ƿ�(intChkTmp).Visible = True
            Else
                GoTo mark01
            End If
        'elseif fld.Type
        Else
mark01:
            intTemp = lblEdit.Count
            Load lblEdit(intTemp)
            Load txtEdit(intTemp)
            
            If fld.Type = adNumeric Then
                '������
                mlng����(intTemp) = 1
            ElseIf fld.Type = adDate Or fld.Type = adDBTimeStamp Or fld.Type = adDBDate Or fld.Type = adDBTime Then
                mlng����(intTemp) = 2
            ElseIf fld.Type = adVarChar Or fld.Type = adLongVarChar Then
                mlng����(intTemp) = 3
            End If
            
            '�����ĸ���ܳ���9
            lblEdit(intTemp).Caption = fld.Name & "(&" & intTemp + lngcboTmp & ")"
            
            '��¼��һЩ�����ֶε����
            Select Case fld.Name
                Case "����"
                    mint���� = intTemp
                Case "����"
                    mint���� = intTemp
                Case "����"
                    mint���� = intTemp
                    mint���볤�� = fld.DefinedSize
            End Select
            lblEdit(intTemp).Tag = fld.Name
            lblEdit(intTemp).Left = 200
            txtEdit(intTemp).Left = lblEdit(intTemp).Left + lblEdit(intTemp).Width + 100
            
            If fld.Type = adVarChar Then
                txtEdit(intTemp).MaxLength = fld.DefinedSize
                txtEdit(intTemp).Width = 300 + fld.DefinedSize * 100
            ElseIf fld.Type = adDate Or fld.Type = adDBTimeStamp Or fld.Type = adDBDate Or fld.Type = adDBTime Then
                txtEdit(intTemp).MaxLength = 10
                txtEdit(intTemp).Width = 300 + fld.Precision * 100
            Else
                txtEdit(intTemp).MaxLength = fld.Precision
                txtEdit(intTemp).Width = 300 + fld.Precision * 100
            End If
            If txtEdit(intTemp).Width > 3000 Then txtEdit(intTemp).Width = 3000
            If chkLog.Width + 200 > sngMaxW Then sngMaxW = chkLog.Width + 200
            If fld.Type = adDate Or fld.Type = adDBTimeStamp Or fld.Type = adDBDate Or fld.Type = adDBTime Then
                txtEdit(intTemp).Text = Format(fld.Value, "yyyy-mm-dd")
            Else
                txtEdit(intTemp).Text = IIf(IsNull(fld.Value), "", fld.Value)
            End If
            txtEdit(intTemp).Top = sngY
            lblEdit(intTemp).Top = txtEdit(intTemp).Top + 75
            sngY = sngY + txtEdit(intTemp).Height + 100
            If txtEdit(intTemp).Width + txtEdit(intTemp).Left > sngMaxW Then sngMaxW = txtEdit(intTemp).Width + txtEdit(intTemp).Left
            lblEdit(intTemp).Visible = True
            txtEdit(intTemp).Visible = True
            
            If fld.Name = "��ɫ" Then
                txtEdit(intTemp).Locked = True
                cmd�ϼ�.Left = txtEdit(intTemp).Left + txtEdit(intTemp).Width
                cmd�ϼ�.Top = txtEdit(intTemp).Top + 10
                cmd�ϼ�.Tag = "��ɫ"
                cmd�ϼ�.Visible = True
                txtEdit(intTemp).Tag = "��ɫ"
                txtEdit(intTemp).Text = IIf(IsNull(fld.Value), 0, fld.Value)
                If txtEdit(intTemp).Text = "" Then txtEdit(intTemp) = 0
                txtEdit(intTemp).ForeColor = txtEdit(intTemp).Text
            End If
            
            '����Tab˳��
            lblEdit(intTemp).TabIndex = (intTemp - 1) * 2
            txtEdit(intTemp).TabIndex = (intTemp - 1) * 2 + 1
            If fld.Name = "�ϼ�" Then
                txtEdit(intTemp).Enabled = False
                If txtEdit(intTemp).Text = "" And str�ϼ� <> "" Then
                    If str�ϼ� <> "oot" Then
                        txtEdit(intTemp).Text = str�ϼ�
                    End If
                End If
                cmd�ϼ�.Left = txtEdit(intTemp).Left + txtEdit(intTemp).Width
                cmd�ϼ�.Top = txtEdit(intTemp).Top + 10
                If cmd�ϼ�.Width + txtEdit(intTemp).Width + txtEdit(intTemp).Left > sngMaxW Then sngMaxW = cmd�ϼ�.Left + cmd�ϼ�.Width
                cmd�ϼ�.Visible = True
                cmd�ϼ�.TabIndex = (intTemp - 1) * 2 + 2
                cmd�ϼ�.Tag = intTemp
            End If
            
        End If
makContinue:
    Next
    
    If chkLog.Tag <> "" Then
        chkLog.Top = sngY
        sngY = sngY + chkLog.Height + 100 '�ѿ�ѡ
        chkLog.TabIndex = intTemp * 2
    End If
    
    If mstr���� = "" Then txtEdit(mint����).Text = zlDatabase.GetMax(mstrOwner & "." & strTable, "����", mint���볤��)
    fraSplit.Top = -500
    fraSplit.Left = sngMaxW + 250
    cmdOK.Left = sngMaxW + 500
    cmdCancel.Left = cmdOK.Left
    
    frmDictEdit.Width = cmdOK.Left + cmdOK.Width + 250
    frmDictEdit.Height = sngY + 500
    'Ϊ����ʾ�꼸����ť����ʹ�������ۡ����ڵĸ߶ȱ�֤��һ����ֵ֮��
    If frmDictEdit.Height < 2300 Then frmDictEdit.Height = 2300
    fraSplit.Height = frmDictEdit.Height + 1000
    
    frmDictEdit.Caption = mstrTable & IIf(intĩ�� = 0, "[����]", "[��Ŀ]")
    frmDictEdit.txtEdit(1).SetFocus
    
    mblnChange = False
    InitEnable
    frmDictEdit.Show vbModal
End Function
Private Sub InitEnable()
Dim intTemp As Integer, rsTemp As ADODB.Recordset
    On Error GoTo ErrH
    For intTemp = 0 To txtEdit.UBound
        Select Case lblEdit(intTemp).Tag
            Case "վ��"
'                If gstrNodeNo <> "-" Then
                    txtEdit(intTemp).Enabled = True
                    txtEdit(intTemp).BackColor = &HFFFFFF
'                Else
'                    txtEdit(intTemp).Enabled = False
'                    txtEdit(intTemp).BackColor = &H80000000
'                End If
        End Select
    Next
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub chkLog_Click()
    mblnChange = True
End Sub

Private Sub chkLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    On Error Resume Next
    If Index = mint���� Then
        txtEdit(mint����).Text = zlCommFun.SpellCode(txtEdit(Index).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If lblEdit(Index).Tag = "����" Then
        zlCommFun.OpenIme True
    ElseIf lblEdit(Index).Tag = "����" Or lblEdit(Index).Tag = "����" Or mlng����(Index) = 1 Then
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    Else
        Select Case lblEdit(Index).Tag
            Case "����"
                If mlng����(Index) = 3 Then
                    If InStr("'", Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    End If
                Else
                    If InStr("0123456789" & Chr(vbKeyBack) & Chr(vbKeyDelete), Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    End If
                End If
            Case "վ��"
                If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
        End Select
    End If
End Sub

'Private Sub SetPathProp(ByVal intVal As Integer)
'    Dim rsTmp As ADODB.Recordset
'    Dim strTmp As String
'    Dim i As Integer
'    strTmp = "select ����,���� from ·��������� order by ����,����"
'    Set rsTmp = zldatabase.OpenSQLRecord(strTmp, Me.Caption)
'    If Not rsTmp.EOF Then
'        For i = 0 To rsTmp.RecordCount - 1
'            cboEdit(intVal).AddItem rsTmp!����
'            cboEdit(intVal).ItemData(i) = rsTmp!����
'            rsTmp.MoveNext
'        Next
'    End If
'End Sub

Private Sub SetSelectProp(ByVal lngVal As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    On Error GoTo errHandle
    If UBound(mvar���ӹ�ϵ) = 2 Then
        strTmp = "select ����,���� from " & mvar���ӹ�ϵ(2) & " order by ����,���� "
        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "��ȡ������������")
        If Not rsTmp.EOF Then
            For i = 0 To rsTmp.RecordCount - 1
                cboEdit(lngVal).AddItem "" & rsTmp!���� & "-" & rsTmp!����
                cboEdit(lngVal).ItemData(i) = rsTmp!����
                rsTmp.MoveNext
            Next
            rsTmp.Close
        End If
    End If
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

