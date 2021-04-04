VERSION 5.00
Begin VB.Form frmPatholSpecimenCfg 
   Caption         =   "����걾����"
   ClientHeight    =   5940
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10455
   Icon            =   "frmPatholSpecimenCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   10455
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   10095
      TabIndex        =   2
      Top             =   5280
      Width           =   10095
      Begin VB.CommandButton cmdExit 
         Caption         =   "��  ��(&E)"
         Height          =   400
         Left            =   8880
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox cbxSpecimenPart 
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
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ  ��(&D)"
         Height          =   400
         Left            =   6240
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "��  ��(&S)"
         Height          =   400
         Left            =   7560
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label labSpecimenPart 
         Caption         =   "�걾��λ��"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   195
         Width           =   975
      End
   End
   Begin VB.Frame framSpecimenCfg 
      Caption         =   "������걾��¼"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7858
         DefaultCols     =   ""
         GridRows        =   201
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
End
Attribute VB_Name = "frmPatholSpecimenCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub AdjustFace()
'�������沼��
    framSpecimenCfg.Left = 120
    framSpecimenCfg.Top = 120
    framSpecimenCfg.Width = Me.Width - 360
    framSpecimenCfg.Height = Me.Height - picControl.Height - 900
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framSpecimenCfg.Width - 240
    ufgData.Height = framSpecimenCfg.Height - 360
    
    
    picControl.Left = 120
    picControl.Top = Me.Height - picControl.Height - 620
    picControl.Width = Me.Width - 360
    
    
    cbxSpecimenPart.Top = 0
    
    labSpecimenPart.Left = 0
    labSpecimenPart.Top = cbxSpecimenPart.Top + 30
    
    cbxSpecimenPart.Left = labSpecimenPart.Left + labSpecimenPart.Width + 60
    
    cmdExit.Left = picControl.Width - cmdSave.Width
    cmdExit.Top = 0
    
    cmdSave.Left = cmdExit.Left - cmdSave.Width - 120
    cmdSave.Top = 0
    
    cmdDel.Left = cmdSave.Left - cmdDel.Width - 120
    cmdDel.Top = 0
    
End Sub



Private Sub InitStudySpecimenList()
    '��������  ���û�������Ϊ500��
    ufgData.GridRows = 501
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.DefaultColNames = gstrSpecimenModuleCols
    
    '��ֹ�Ҽ������б����ô���
    ufgData.IsEjectConfig = False
    '��ʼ��������걾�б�
    ufgData.ColNames = gstrSpecimenModuleCols
    ufgData.ColConvertFormat = gstrSpecimenModuleConvertFormat
End Sub


Private Sub cbxSpecimenPart_Click()
'�����ײ���Ϣ
On Error GoTo ErrHandle
    Dim strSQL As String
    
    If cbxSpecimenPart.Text = "" Then
        strSQL = "select ID,�걾����,�걾��λ,�걾����,Ĭ�ϱ걾��,Ĭ����Ƭ��,����,��ע from ������걾 order by �걾��λ,�걾����"
    Else
        strSQL = "select ID,�걾����,�걾��λ,�걾����,Ĭ�ϱ걾��,Ĭ����Ƭ��,����,��ע from ������걾 where �걾��λ=[1] order by �걾��λ,�걾����"
    End If
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbxSpecimenPart.Text)
    
    Call ufgData.RefreshData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdDel_Click()
'ɾ��������걾
On Error GoTo ErrHandle
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ����ɾ���Ĳ�����걾��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫɾ���ò�����걾������", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call ufgData.DelCurRow
    
    '����ɾ��������
    Call SaveStudySpeciments(True)
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrHandle
    Call Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cmdSave_Click()
On Error GoTo ErrHandle
    Dim blnValid As Boolean
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "��⵽�������б��д�����Ч���ݣ���ȷ����������Ƿ���ȷ������¼�룬����ɫ����ǵĵ�Ԫ��Ϊ��¼���ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�����ײ���Ϣ
    Call SaveStudySpeciments
    
    Call ConfigInput
    
    Call MsgBoxD(Me, "�����ѱ���ɹ���", vbOKOnly, Me.Caption)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitStudySpecimenList
    Call LoadStudySpecimenMoudleData
    
    Call ConfigInput
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ConfigInput()
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strSpecimenParts As String
    
    '��ȡ�Ѿ����ڵı걾��λ
    strSQL = "select distinct(�걾��λ) as �걾��λ from ������걾 order by �걾��λ"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    ufgData.ComboxListFormat(ufgData.GetColIndex(gstrSpecimenModule_�걾��λ)) = ""
    cbxSpecimenPart.Clear
     
    If rsData.RecordCount > 0 Then

        Call cbxSpecimenPart.AddItem("")
        
        strSpecimenParts = "|"
        
        While Not rsData.EOF
            If Nvl(rsData!�걾��λ) <> "" Then
                
                If strSpecimenParts <> "|" Then strSpecimenParts = strSpecimenParts & "|"
                
                strSpecimenParts = strSpecimenParts & Nvl(rsData!�걾��λ)
                Call cbxSpecimenPart.AddItem(Nvl(rsData!�걾��λ))
            
            End If
            rsData.MoveNext
        Wend
        
        If strSpecimenParts = "|" Then Exit Sub
        ufgData.ComboxListFormat(ufgData.GetColIndex(gstrSpecimenModule_�걾��λ)) = strSpecimenParts
    End If
End Sub


Private Sub LoadStudySpecimenMoudleData()
'���벡����걾ģ������
    Dim strSQL As String
    
    strSQL = "select ID,�걾����,�걾��λ,�걾����,Ĭ�ϱ걾��,Ĭ����Ƭ��,����,��ע from ������걾 order by �걾��λ,�걾����"
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call ufgData.RefreshData
End Sub


Private Sub SaveStudySpeciments(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'blnIsSaveOnlyDel:�Ƿ񱣴��ɾ��������

'���没����걾����
    Dim i As Long
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    For i = 1 To ufgData.GridRows - 1
        Select Case ufgData.RowState(i)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Add)
                
                strSQL = "select ZL_����걾����_����([1],[2],[3],[4],[5],[6],[7]) as ����ֵ from dual"
                Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                        ufgData.Text(i, gstrSpecimenModule_�걾����), _
                                                        ufgData.Text(i, gstrSpecimenModule_�걾��λ), _
                                                        Val(ufgData.Text(i, gstrSpecimenModule_�걾����)), _
                                                        Val(ufgData.Text(i, gstrSpecimenModule_Ĭ�ϱ걾��)), _
                                                        Val(ufgData.Text(i, gstrSpecimenModule_Ĭ����Ƭ��)), _
                                                        ufgData.Text(i, gstrSpecimenModule_����), _
                                                        ufgData.Text(i, gstrSpecimenModule_��ע))
                                                        
                If rsData.RecordCount <= 0 Then
                    Call err.Raise(0, "SaveStudySpeciments", "δ�ɹ���ȡ������Ĳ�����걾ID,����ʧ�ܡ�")
                    Exit Sub
                End If
                
                
                Call ufgData.SyncText(i, gstrSpecimenModule_ID, rsData!����ֵ)
				
				ufgData.RowState(i) = TDataRowState.Normal
                                                        
            Case TDataRowState.Del
                strSQL = "ZL_����걾����_ɾ��(" & Val(ufgData.KeyValue(i)) & ")"
                
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
				
				ufgData.RowState(i) = TDataRowState.Normal
                
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Update)
                strSQL = "ZL_����걾����_����(" & Val(ufgData.KeyValue(i)) & ",'" & _
                                                ufgData.Text(i, gstrSpecimenModule_�걾����) & "','" & _
                                                ufgData.Text(i, gstrSpecimenModule_�걾��λ) & "'," & _
                                                Val(ufgData.Text(i, gstrSpecimenModule_�걾����)) & ",'" & _
                                                Val(ufgData.Text(i, gstrSpecimenModule_Ĭ�ϱ걾��)) & "'," & _
                                                Val(ufgData.Text(i, gstrSpecimenModule_Ĭ����Ƭ��)) & ",'" & _
                                                ufgData.Text(i, gstrSpecimenModule_����) & "','" & _
                                                ufgData.Text(i, gstrSpecimenModule_��ע) & "')"
                                                
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
				
				ufgData.RowState(i) = TDataRowState.Normal

        End Select
        
    Next i
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    
    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        
        Exit Sub
    End If
        
    
    '���δ¼��걾���ƣ�����ʾ����ɫ
    iCol = ufgData.GetColIndex(gstrSpecimenModule_�걾����)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrSpecimenModule_�걾����) = "", ufgData.ErrCellColor, ufgData.BackColor)
           
    
    
    '���Ĭ����Ƭ��С��1������ʾ����ɫ
    iCol = ufgData.GetColIndex(gstrSpecimenModule_Ĭ����Ƭ��)
    
    ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrSpecimenModule_Ĭ����Ƭ��)) < 1, ufgData.ErrCellColor, ufgData.BackColor)
    
    
    '���ü��걾����
    If ufgData.Text(Row, gstrSpecimenModule_�걾����) <> "" Then
        If ufgData.Text(Row, gstrSpecimenModule_����) = "" Then ufgData.Text(Row, gstrSpecimenModule_����) = zlCommFun.SpellCode(ufgData.Text(Row, gstrSpecimenModule_�걾����))
    End If
End Sub



Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strComboboxText As String
    
    If Row > 0 Then
        '�Զ���д�걾��λ
        If ufgData.Text(Row, gstrSpecimenModule_�걾��λ) = "" Then
            If Row - 1 > 0 Then
                If ufgData.Text(Row - 1, gstrSpecimenModule_�걾��λ) <> "" Then
                    ufgData.Text(Row, gstrSpecimenModule_�걾��λ) = ufgData.Text(Row - 1, gstrSpecimenModule_�걾��λ)
                End If
            End If
            
            If ufgData.Text(Row, gstrSpecimenModule_�걾��λ) = "" Then
                strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrSpecimenModule_�걾��λ))
                
                If strComboboxText <> "" Then
                    If InStr(strComboboxText, ";") > 0 Then
                        strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, ";") - 1)
                    End If
                    ufgData.Text(Row, gstrSpecimenModule_�걾��λ) = Mid(strComboboxText, InStr(strComboboxText, "#") + 1, 255)
                    
                End If
            End If
        End If
        
        '�Զ���д�걾����
        If ufgData.Text(Row, gstrSpecimenModule_�걾����) = "" Then
                If Row - 1 > 0 Then
                    If ufgData.Text(Row - 1, gstrSpecimenModule_�걾����) <> "" Then
                        ufgData.Text(Row, gstrSpecimenModule_�걾����) = ufgData.Text(Row - 1, gstrSpecimenModule_�걾����)
                    End If
                End If
                
                If ufgData.Text(Row, gstrSpecimenModule_�걾����) = "" Then
                    strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrSpecimenModule_�걾����))
                    
                    If strComboboxText <> "" Then
                        If InStr(strComboboxText, "|") > 0 Then
                            strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, "|") - 1)
                        End If
                        ufgData.Text(Row, gstrSpecimenModule_�걾����) = strComboboxText
                        
                    End If
                End If
        End If
                
        '����Ĭ����Ƭ��
        If Val(ufgData.Text(Row, gstrSpecimenModule_Ĭ����Ƭ��)) <= 0 Then ufgData.Text(Row, gstrSpecimenModule_Ĭ����Ƭ��) = "1"
        
    End If
End Sub
