VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{099B2A6C-9CCE-43CF-AEF0-C526C98F4B7F}#1.1#0"; "ZLRICHEDITOR.OCX"
Begin VB.Form frmDockEPRContent 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "�����ļ����"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin zlRichEditor.Editor edtThis 
      Height          =   2580
      Left            =   630
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4551
      WithViewButtonas=   0   'False
      ShowRuler       =   0   'False
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmDockEPRContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event DblClick()                                                 '����˫�������¼�

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mlngRecordId As Long        '������¼ID

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
    End Select
End Sub

Private Sub Form_Load()
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    mlngRecordId = -1
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.edtThis
        .Left = Me.ScaleLeft + 120: .Width = Me.ScaleWidth - 2 * .Left
        .Top = Me.ScaleTop + 120: .Height = Me.ScaleHeight - 2 * .Top
    End With
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Me.edtThis.Copy
    End Select
End Sub

Private Sub edtThis_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    RaiseEvent DblClick
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, x As Single, y As Single)
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
        Popup.ShowPopup
    End With
End Sub

'-----------------------------------------------------
'���幫������
'-----------------------------------------------------

Public Sub zlRefresh(ByVal lngRecordId As Long, Optional ByVal blnPrivacyProtect As Boolean)
    '���ܣ�ˢ�²�����ʾ���ݣ�
    '������lngRecordId�����Ӳ�����¼ID��blnPrivacyProtect���Ƿ�������˽����
    Dim mstrPrivs As String, blnPrivacy As Boolean, Elements As New cEPRElements
    Dim RS As New ADODB.Recordset, lngKey As Long
    If blnPrivacyProtect = True Then
        mstrPrivs = ";" & GetPrivFunc(glngSys, 1070) & ";"
        blnPrivacy = InStr(mstrPrivs, ";������˽����;") = 0     '������˽��Ŀ
    End If
    
    Dim strTemp As String, strZipFile As String
'    If mlngRecordId = lngRecordId Then Exit Sub
    mlngRecordId = lngRecordId
    Me.edtThis.Freeze
    Me.edtThis.ReadOnly = False
    Me.edtThis.NewDoc
    strZipFile = zlBlobRead(5, lngRecordId)
    If gobjFSO.FileExists(strZipFile) Then
        strTemp = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strTemp) Then
            '���ļ�
            Me.edtThis.OpenDoc strTemp
            '�����滻��Ŀ
            If blnPrivacy Then
                '��ȡ���е�Ҫ��
                gstrSQL = "Select A.ID,A.������ From ���Ӳ������� A, ��˽������Ŀ B,����������Ŀ C " & _
                    "Where A.�������� = 4 And A.�滻�� = 1 And A.�ļ�id = [1] And A.������� > 0 and B.��Ŀid = C.ID And A.Ҫ������ =C.������ And C.�滻�� = 1 "
                Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
                If Not RS.EOF Then
                    Do While Not RS.EOF
                        lngKey = Elements.Add(NVL(RS("������"), 0))
                        Elements("K" & lngKey).GetElementFromDB cprET_�������༭, RS("ID"), True, "���Ӳ�������"
                        '�滻Ҫ������
                        Elements("K" & lngKey).�����ı� = String(Len(Elements("K" & lngKey).�����ı�), "*")
                        Elements("K" & lngKey).Refresh Me.edtThis
                        RS.MoveNext
                    Loop
                End If
                RS.Close
            End If
            gobjFSO.DeleteFile strTemp, True
        End If
        gobjFSO.DeleteFile strZipFile, True
        Me.edtThis.SelStart = 0
    End If
    If lngRecordId > 0 Then
        '����ҳ���ʽ
        Dim mEPRFileInfo As New cEPRFileDefineInfo
        gstrSQL = "Select c.ID, a.��ʽ From   ����ҳ���ʽ a, �����ļ��б� b, ���Ӳ�����¼ c " & _
                " Where  c.�ļ�id = b.id And a.���� = b.���� And a.��� = b.ҳ�� And c.ID = [1]"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
        If Not RS.EOF Then
            mEPRFileInfo.��ʽ = zlCommFun.NVL(RS("��ʽ").Value)
            mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.��ʽ
            Me.edtThis.ResetWYSIWYG
        End If
        Set mEPRFileInfo = Nothing
    End If
    Me.edtThis.UnFreeze
    edtThis.RefreshTargetDC
    Me.edtThis.ReadOnly = True
End Sub
