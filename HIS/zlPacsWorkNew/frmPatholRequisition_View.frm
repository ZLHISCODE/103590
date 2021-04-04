VERSION 5.00
Begin VB.Form frmPatholRequisition_View 
   Caption         =   "����鿴"
   ClientHeight    =   6390
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   9075
   Icon            =   "frmPatholRequisition_View.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9075
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   7680
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame framRequest 
      Caption         =   "�����¼"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9128
         GridRows        =   21
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
End
Attribute VB_Name = "frmPatholRequisition_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnMoved As Boolean


Public Sub ShowRequestViewWind(ByVal lngPatholAdviceId As Long, ByVal lngRequestType As Long, _
    ByVal blnMoved As Boolean, owner As Form)
'��ʾ����鿴����
    mblnMoved = blnMoved
    
    Call LoadRequestViewData(lngPatholAdviceId, lngRequestType)
    
    Select Case lngRequestType
        Case 0
            Me.Caption = "��������鿴"
        Case 1
            Me.Caption = "��Ⱦ����鿴"
        Case 2
            Me.Caption = "��������鿴"
        Case 3
            Me.Caption = "��Ƭ����鿴"
        Case 4
            Me.Caption = "ȡ������鿴"
    End Select
    
    Call Me.Show(1, owner)
End Sub



Private Sub InitRequestList()
'��ʼ������鿴�б�
     Dim strTemp As String
     

    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("����鿴�б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrRequisitionViewCols
     
    If strTemp = "" Then
        ufgData.ColNames = gstrRequisitionViewCols
    Else
        ufgData.ColNames = strTemp
    End If
         '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrRequisitionConvertFormat
End Sub



Private Sub LoadRequestViewData(ByVal lngPatholAdviceId As Long, ByVal lngRequestType As Long)
'����������Ϣ
    Dim strSql As String
    
    strSql = "select ����ID,������,����ʱ��,����ϸĿ,����״̬,��������,����״̬,���ʱ�� from ����������Ϣ where ����ҽ��ID=[1] and ��������=[2] order by ����ʱ��"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId, lngRequestType)
    
    Call ufgData.RefreshData
End Sub


Private Sub AdjustFace()
    framRequest.Left = 120
    framRequest.Top = 120
    framRequest.Width = Me.Width - 360
    framRequest.Height = Me.Height - cmdSure.Height - 900
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framRequest.Width - 240
    ufgData.Height = framRequest.Height - 360
    
    cmdSure.Left = Me.Width - cmdSure.Width - 240
    cmdSure.Top = Me.Height - cmdSure.Height - 620
End Sub


Private Sub cmdSure_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitRequestList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    '�رմ���ʱ�����б�����
     zlDatabase.SetPara "����鿴�б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
     
End Sub
