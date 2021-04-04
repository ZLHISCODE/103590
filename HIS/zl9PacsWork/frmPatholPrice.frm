VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPatholPrice 
   Caption         =   "��鲹��"
   ClientHeight    =   7080
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   9195
   Icon            =   "frmPatholPrice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9195
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picRequest 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6855
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.Frame framDetails 
         Caption         =   "������ϸ"
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   4575
         Begin zl9PACSWork.ucFlexGrid ufgContext 
            Height          =   2535
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4471
            DefaultCols     =   ""
            GridRows        =   21
            IsKeepRows      =   0   'False
            BackColor       =   12648447
            IsEnterNextCell =   0   'False
            IsBtnNextCell   =   0   'False
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            Editable        =   1
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
      Begin VB.Frame framRequest 
         Caption         =   "�����¼"
         Height          =   3015
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   4575
         Begin zl9PACSWork.ucFlexGrid ufgRequest 
            Height          =   2655
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   4335
            _ExtentX        =   16325
            _ExtentY        =   3413
            DefaultCols     =   ""
            GridRows        =   21
            IsKeepRows      =   0   'False
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            Editable        =   0
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
      Begin VB.PictureBox picControl 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -240
         ScaleHeight     =   495
         ScaleWidth      =   4935
         TabIndex        =   1
         Top             =   6360
         Width           =   4935
         Begin VB.CommandButton cmdAlreadyPrice 
            Caption         =   "��ɲ���(&F)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   240
            TabIndex        =   9
            Top             =   0
            Width           =   4575
         End
         Begin VB.CommandButton cmdTempPrice 
            Caption         =   "�� �� (&T)"
            Enabled         =   0   'False
            Height          =   400
            Left            =   240
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdBill 
            Caption         =   "�� ��(&M)"
            Enabled         =   0   'False
            Height          =   400
            Left            =   240
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdAccept 
            Caption         =   "�� ��(&R)"
            Enabled         =   0   'False
            Height          =   400
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkAutoExecute 
            Caption         =   "���Ѻ�����Զ�ִ��"
            Height          =   255
            Left            =   2760
            TabIndex        =   10
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   5400
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mlngModule As Long
Private mstrPrivs As String
Private mlngCurDepartmentId As Long
Private mobjOwner As Object

Private mlngCurAdviceId As Long
Private mlngSendNo As Long
Private mblnMoved As Boolean
Private mblnReadOnly As Boolean


Private mlngRequestType As Long

Private mlngCurRequestId As Long
Private mblnButtonEvent As Boolean

Private mrecStudyInf As TStudyStateInf

Private mobjExpense As zlPublicExpense.clsDockExpense        '���ö���


Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
'��ʼ��ģ�����
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngCurDepartmentId = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner
End Sub

Public Sub zlRefresh(ByVal lngCurDepartmentId As Long, lngAdviceID As Long, ByVal lngSendNO As Long, ByVal blnMoved As Boolean)
    
On Error GoTo errHandle
    If lngAdviceID <= 0 Then
        Call ConfigPriceFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If
    
    mlngCurAdviceId = lngAdviceID
    mblnMoved = blnMoved
    mlngSendNo = lngSendNO
    mlngCurDepartmentId = lngCurDepartmentId
    mblnReadOnly = blnMoved
    
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
    
    
    If mrecStudyInf.strPatholNumber = "" Then
        Call RefreshPrice(lngCurDepartmentId, lngAdviceID, lngSendNO, blnMoved)
        
        Call ConfigPriceFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        
'        If Not (mobjOwner Is Nothing) Then
'            Call MsgBoxD(Me, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�", vbOKOnly, Me.Caption)
'        End If
        
        Exit Sub
    End If
    
    
    '��ȡ������Ϣ
    Call LoadRequestInf(mrecStudyInf.lngPatholAdviceId)
    
    '����������ϸ
    Call ufgRequest_OnClick
    
    Call ConfigPriceFace(True)

    
    Call ConfigPopedom(mblnReadOnly)
    
'    '���ò�����Դ����
'    Call ConfigPatientSource(lngAdviceID)
    
    Call RefreshPrice(lngCurDepartmentId, lngAdviceID, lngSendNO, blnMoved)
    
    If ufgRequest.ShowingDataRowCount > 0 Then
        Call ufgRequest.LocateRow(1)
        Call ConfigPriceState(ufgRequest.Text(1, gstrRequisition_����״̬) = "�貹��")
    Else
        Call ConfigPriceState(False)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ConfigPatientSource(ByVal lngAdviceID As Long)
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ������Դ from ����ҽ����¼ where ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    If Nvl(rsData!������Դ) = 2 Then
        cmdAccept.Enabled = False
    End If

End Sub

Private Sub LoadRequestInf(ByVal lngPatholAdviceId As Long)
'����������Ϣ
    Dim strSql As String
    
    strSql = "select ����ID,������,��������,����״̬,����ϸĿ,����ʱ��,����״̬,��������,���ʱ�� " & _
        " from ����������Ϣ where ����ҽ��ID=[1] order by ��������,����ʱ��"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgRequest.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    Call ufgRequest.RefreshData
End Sub



Private Sub ConfigPriceFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'���ò��ѽ���

    cmdAccept.Enabled = blnIsValid
    cmdBill.Enabled = blnIsValid
    cmdTempPrice.Enabled = blnIsValid
    cmdAlreadyPrice.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgRequest.CloseHintInf
        Call ufgContext.CloseHintInf
    Else
        Call ufgRequest.ShowHintInf(strHintInf)
        Call ufgContext.ShowHintInf(strHintInf)
    End If
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'����Ȩ��
    Dim blnIsAllowPrice As Boolean
    
    blnIsAllowPrice = IIf(GetInsidePrivs(pҽ�����ѹ���, True) <> "", True, False)
    
    
    cmdAccept.Enabled = blnIsAllowPrice And Not blnIsReadOnly
    cmdBill.Enabled = blnIsAllowPrice And Not blnIsReadOnly
    cmdTempPrice.Enabled = blnIsAllowPrice And Not blnIsReadOnly
    
    
    ufgRequest.ReadOnly = blnIsReadOnly
    ufgContext.ReadOnly = blnIsReadOnly
End Sub


Private Sub InitFace()
'��ʼ�����沼��
    Dim Pane1 As Pane, Pane2 As Pane

    If Not mobjExpense Is Nothing Then
        With dkpMain
            .CloseAll
            .Options.HideClient = True
            .Options.UseSplitterTracker = False 'ʵʱ�϶�
            .Options.ThemedFloatingFrames = True
            .Options.AlphaDockingContext = True
        End With
    
        Set Pane1 = dkpMain.CreatePane(1, 0, Round(Me.Height / 2), DockLeftOf, Nothing)
        Pane1.Title = "�����¼"
        Pane1.Handle = picRequest.hWnd
        Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Pane1.MinTrackSize.Width = 50
        Pane1.MinTrackSize.Height = 50
    
        Set Pane2 = dkpMain.CreatePane(2, 0, Round(Me.Height / 2), DockRightOf, Pane1)
        Pane2.Title = "���ü�¼"
        Pane2.Handle = mobjExpense.zlGetForm.hWnd
        Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Pane2.MinTrackSize.Width = 50
        Pane2.MinTrackSize.Height = 50
    Else
        picRequest.Width = Me.ScaleWidth - 240
        picRequest.Height = Me.ScaleHeight
    End If
End Sub

Private Sub cmdAccept_Click()
'On Error GoTo errHandle
'
'    mblnButtonEvent = True
'
'
'    '�շѵ���
'    Call mobjExpense.zlExecuteCommandBars1(1)
'
'    mblnButtonEvent = False
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdAlreadyPrice_Click()
On Error GoTo errHandle
    Dim i As Long
    Dim strSql As String
    Dim lngRequestId As String
    
    If ufgRequest.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "û����Ҫִ�в��ѵļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    'ִ�в���
    For i = 1 To ufgRequest.GridRows - 1
        If ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_����״̬) = "�貹��" Then
            lngRequestId = Val(ufgRequest.KeyValue(ufgRequest.SelectionRow))
    
            strSql = "zl_��������_����(" & lngRequestId & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
            '���������б�Ĳ���״̬
            ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_����״̬) = "�Ѳ���"
        End If
    Next i

    
    '���·���״̬Ϊ����ִ��
    If chkAutoExecute.value <> 0 Then
        Call ExecuteStudyMoney
    End If

        
    Call MsgBoxD(Me, "����ɲ��Ѳ�����", vbOKOnly, Me.Caption)
    
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



'����ִ��
Private Sub ExecuteStudyMoney()
    On Error GoTo errHandle
      
    
    gstrSQL = "Zl_Ӱ�����ִ��(" & mlngCurAdviceId & "," & mlngSendNo & ",2,Null,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCurDepartmentId & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "����ִ��"
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



Private Sub cmdBill_Click()
'On Error GoTo errHandle
'    mblnButtonEvent = True
'
'    '���˵���
'    Call mobjExpense.zlExecuteCommandBars1(2)
'
'    mblnButtonEvent = False
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdTempPrice_Click()
'On Error GoTo errHandle
'    mblnButtonEvent = True
'
'    '��ķ���
'    Call mobjExpense.zlExecuteCommandBars1(3)
'
'    mblnButtonEvent = False
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
    Dim objTmp As zlPublicExpense.clsPublicExpense
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    mblnButtonEvent = False
    mlngCurRequestId = -1
    
    If mlngModule > -1 Then
        Set objTmp = New zlPublicExpense.clsPublicExpense
        Call objTmp.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
        Set mobjExpense = objTmp.zlDockExpense
    End If
    
    Call InitFace
    
        '��ʼ�������б�
    Call InitRequisitionList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte, ByVal bytSize As Byte)
On Error GoTo errHandle

    Call ReSetFormFontSize(bytFontSize)
    
    If Not mobjExpense Is Nothing Then
        Call mobjExpense.SetFontSize(bytSize)
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'����:�������ù���վ����������С
    
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    
    Me.FontSize = bytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Frame")
            objCtrl.Font.Size = bytFontSize
        Case UCase("Label")
            objCtrl.FontSize = bytFontSize
            objCtrl.Height = TextHeight("��") + 20
        Case UCase("ucFlexGrid")
            objCtrl.DataGrid.Cell(flexcpFontSize, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = bytFontSize
            objCtrl.DataGrid.FontSize = bytFontSize
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = bytFontSize
        End Select
    Next
    
End Sub

Private Sub InitRequisitionList()
'��ʼ�������б�
    Dim strTemp As String
    

    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("��������б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgRequest.DefaultColNames = gstrRequisitionCols
     
    If strTemp = "" Then
        ufgRequest.ColNames = gstrRequisitionCols
    Else
        ufgRequest.ColNames = strTemp
    End If
        '��������
    ufgRequest.GridRows = glngStandardRowCount
    '�����и�
    ufgRequest.RowHeightMin = glngStandardRowHeight
    ufgRequest.ColConvertFormat = gstrRequisitionConvertFormat
End Sub

Private Sub InitRequestContextList(ByVal lngRequestType As Long)
'��ʼ��������Ŀ��ϸ�б�
    Dim strTemp As String
    

    
    mlngRequestType = lngRequestType
    
    Select Case lngRequestType
        Case 0, 1, 2
        
            '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
            strTemp = zlDatabase.GetPara("�ؼ������б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
            ufgContext.DefaultColNames = gstrRequest_SpeExam_Cols
                        
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_SpeExam_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
            
            ufgContext.ColConvertFormat = gstrRequest_SpeExamConvertFormat
            
        Case 3
            
            '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
            strTemp = zlDatabase.GetPara("��Ƭ�����б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
            ufgContext.DefaultColNames = gstrRequest_Slices_Cols
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_Slices_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
            
            ufgContext.ColConvertFormat = gstrRequest_SlicesConvertFormat
        Case 4, 5
            
            '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
            strTemp = zlDatabase.GetPara("��ȡ�����б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
            ufgContext.DefaultColNames = gstrRequest_Material_Cols
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_Material_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
            
            ufgContext.ColConvertFormat = gstrRequest_MaterialConvertFormat
    End Select
        '��������
    ufgContext.GridRows = glngStandardRowCount
    '�����и�
    ufgContext.RowHeightMin = glngStandardRowHeight
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If mlngModule = -1 Then
        picRequest.Width = Me.ScaleWidth - 240
        picRequest.Height = Me.ScaleHeight
    End If
err.Clear
End Sub

Private Sub ufgContext_OnColFormartChange()
    '����ı�ʱ�����б�����
    
    Select Case mlngRequestType
        Case 0, 1, 2
        
            zlDatabase.SetPara "�ؼ������б�����", ufgContext.GetColsString(ufgContext), glngSys, G_LNG_PATHOLSYS_NUM
            
        Case 3
        
            zlDatabase.SetPara "��Ƭ�����б�����", ufgContext.GetColsString(ufgContext), glngSys, G_LNG_PATHOLSYS_NUM
           
        Case 4, 5
            
            zlDatabase.SetPara "��ȡ�����б�����", ufgContext.GetColsString(ufgContext), glngSys, G_LNG_PATHOLSYS_NUM
            
    End Select

End Sub

Private Sub ufgRequest_OnColFormartChange()
'����ı�ʱ�����б�����
     zlDatabase.SetPara "��������б�����", ufgRequest.GetColsString(ufgRequest), glngSys, G_LNG_PATHOLSYS_NUM
     
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjExpense Is Nothing Then
        Unload mobjExpense.zlGetForm
    End If
    
    Set mobjExpense = Nothing
End Sub



'Private Sub mobjExpense_OnPriceEvent(ByVal lngAdviceID As Long, ByVal lngPriceType As Long)
'    Dim strSQL As String
'
'    If mblnButtonEvent And mlngCurRequestId > 0 Then
'        strSQL = "zl_��������_����(" & mlngCurRequestId & ")"
'        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
'
'        '���������б�Ĳ���״̬
'        ufgRequest.Text(ufgRequest.SelectRowIndex, gstrRequisition_����״̬) = "�Ѳ���"
'
'        Call ConfigPriceState(ufgRequest.Text(ufgRequest.SelectRowIndex, gstrRequisition_����״̬) = "�貹��")
'    End If
'End Sub

Private Sub picRequest_Resize()
On Error Resume Next
     Call AdjustFace
End Sub



Private Sub AdjustFace()
    Dim lngAvgHeight As Long
    
    lngAvgHeight = Fix((picRequest.Height - picControl.Height) / 2)
    
    framRequest.Left = 0
    framRequest.Top = 0
    framRequest.Width = picRequest.Width
    framRequest.Height = lngAvgHeight - 120
    
    ufgRequest.Left = 120
    ufgRequest.Top = 240
    ufgRequest.Width = framRequest.Width - 240
    ufgRequest.Height = framRequest.Height - 360
    
    
    framDetails.Left = 0
    framDetails.Top = framRequest.Top + framRequest.Height + 60
    framDetails.Width = picRequest.Width
    framDetails.Height = lngAvgHeight - 120
    
    
    ufgContext.Left = 120
    ufgContext.Top = 240
    ufgContext.Width = framDetails.Width - 240
    ufgContext.Height = framDetails.Height - 360
    
    
    picControl.Left = 0
    picControl.Top = framDetails.Top + framDetails.Height + 60
    picControl.Width = framRequest.Width
    
    
    cmdAlreadyPrice.Left = 0 'picControl.Width - cmdAlreadyPrice.Width
    cmdAlreadyPrice.Top = 60
    cmdAlreadyPrice.Width = framDetails.Width
    
'    chkAutoExecute.Left = 0 'cmdAlreadyPrice.Left + cmdAlreadyPrice.Width + 120
'    chkAutoExecute.Top = cmdAlreadyPrice.Top
    
    
'    cmdTempPrice.Left = picControl.Width - cmdTempPrice.Width
'    cmdTempPrice.Top = 0
'
'    cmdBill.Left = cmdTempPrice.Left - cmdBill.Width - 120
'    cmdBill.Top = 0
'
'    cmdAccept.Left = cmdBill.Left - cmdAccept.Width - 120
'    cmdAccept.Top = 0
    

    
End Sub

Private Sub ConfigPriceState(ByVal blnIsPrice As Boolean)
'���ò��Ѱ�ť״̬
    cmdAccept.Enabled = blnIsPrice
    cmdBill.Enabled = blnIsPrice
    cmdTempPrice.Enabled = blnIsPrice
    cmdAlreadyPrice.Enabled = blnIsPrice
End Sub



Private Sub ufgRequest_OnClick()
'��ȡ��������
On Error GoTo errHandle
    Dim strRequestType As String
    
    mlngCurRequestId = -1
    
    '���������Ŀ��ϸ
    Call ufgContext.ClearListData
    Call ConfigPriceState(False)
    
    If Not ufgRequest.IsSelectionRow Then Exit Sub
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then Exit Sub
    
    Call ConfigPriceState(ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_����״̬) = "�貹��")
    
    strRequestType = ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_��������)
    mlngCurRequestId = Val(ufgRequest.KeyValue(ufgRequest.SelectionRow))
    
    Select Case strRequestType
        Case "�����黯", "���Ӳ���", "����Ⱦɫ"
        
            Call InitRequestContextList(0)
            
            '��ȡ�ؼ���Ŀ��ϸ
            Call LoadSpeExamRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))
            
        Case "����Ƭ", "����", "����", "����", "��Ƭ"
            
            Call InitRequestContextList(3)
             
            '��ȡ��Ƭ��Ŀ��ϸ
            Call LoadSlicesRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))

        Case "��ȡ��", "��ȡ��"
            
            Call InitRequestContextList(4)
            
            '��ȡȡ����Ŀ��ϸ
            Call LoadSupMaterialRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))
            
    End Select
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub LoadSlicesRequestContext(ByVal lngRequestId As Long)
'��ȡ��Ƭ��������
    Dim strSql As String
    
    strSql = "select a.ID,a.�Ŀ�ID,b.���,b.�걾����,a.��Ƭ����,a.��Ƭ��ʽ,a.��Ƭ��,a.��ǰ״̬,a.��Ƭʱ��,a.��Ƭ�� " & _
            " from ������Ƭ��Ϣ a, ����ȡ����Ϣ b " & _
            " where a.�Ŀ�id=b.�Ŀ�id and a.����id=[1] order by a.��ǰ״̬, b.�걾����,a.�Ŀ�ID"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub



Private Sub LoadSpeExamRequestContext(ByVal lngRequestId As Long)
'��ȡ�ؼ���������
    Dim strSql As String
    
    strSql = "select a.ID,a.�Ŀ�ID,b.���,b.�걾����,c.����ID, b.�걾����,c.��������,a.��������,a.��ǰ״̬,a.��Ŀ���,a.���ʱ��,a.�ؼ�ҽʦ " & _
                " from �����ؼ���Ϣ a, ����ȡ����Ϣ b, ��������Ϣ c " & _
                " where a.�Ŀ�id = b.�Ŀ�id and a.����id=c.����id and a.����id=[1] order by a.��������, a.�Ŀ�ID, c.��������"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub


Private Sub LoadSupMaterialRequestContext(ByVal lngRequestId As Long)
'��ȡȡ�ĵ��������
    Dim strSql As String
    
    strSql = "select �Ŀ�ID,���,�걾����,�걾��,������,ȡ��ʱ��,��ȡҽʦ,��ȡҽʦ,��¼ҽʦ " & _
            " from ����ȡ����Ϣ where  ����id=[1] order by ȡ��ʱ��"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub







Public Sub zlExecuteCommandBars(Control As Object)
    If mobjExpense Is Nothing Then Exit Sub
    
    Call mobjExpense.zlExecuteCommandBars(Control)
End Sub

Public Sub zlDefCommandBars(frmParent As Object, CommandBars As Object)
    If mobjExpense Is Nothing Then Exit Sub
    
    Call mobjExpense.zlDefCommandBars(frmParent, CommandBars)
End Sub


Public Sub zlPopupCommandBars(CommandBar As Object)
    If mobjExpense Is Nothing Then Exit Sub
    
    Call mobjExpense.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlUpdateCommandBars(Control As Object)
    If mobjExpense Is Nothing Then Exit Sub
    
    Call mobjExpense.zlUpdateCommandBars(Control)
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = Me
End Function

Private Sub RefreshPrice(lng����ID As Long, lngҽ��ID As Long, lng���ͺ� As Long, _
    Optional ByVal blnMoved As Boolean, Optional ByVal bln����ִ�� As Boolean)
    If mobjExpense Is Nothing Then Exit Sub
    
'    Call mobjExpense.zlRefresh(lng����ID, lngҽ��ID, lng���ͺ�, blnMoved, bln����ִ��)
    Call mobjExpense.zlRefresh(lng����ID, lngҽ��ID & ":" & lng���ͺ� & ":" & IIf(bln����ִ�� = True, 1, 0), blnMoved)
End Sub

