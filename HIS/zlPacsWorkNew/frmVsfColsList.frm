VERSION 5.00
Begin VB.Form frmVsfColsList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����б�����"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVsfColsList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ListBox lstVsfColsName 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      ItemData        =   "frmVsfColsList.frx":6852
      Left            =   120
      List            =   "frmVsfColsList.frx":6859
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "�ָ�Ĭ��(&D)"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      TabIndex        =   1
      Top             =   3960
      Width           =   820
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&S)"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   0
      Top             =   3960
      Width           =   825
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ��Ҫ��ʾ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmVsfColsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjOwner As Object
Private mblnChange As Boolean

Private mColConfigOld() As clsQueryPar.TColConfig
Private mColConfigNew() As clsQueryPar.TColConfig


Public Function ShowVsfColsListWindow(ByRef ColConfig() As clsQueryPar.TColConfig, ByVal frmParent As Object) As Boolean
'�������б��岢����Ĭ����ʾ����  StrDefaultCfg
    mColConfigNew = ColConfig
    mColConfigOld = ColConfig
    
    mblnChange = False

    Set mobjOwner = frmParent
    
    '���ü���Ĭ����ʾ����
    Call LoadColsList
    
    '���ش���
    Me.Move frmParent.Left + (frmParent.Width - Me.Width) / 2, frmParent.Top + (frmParent.Height - Me.Height) / 2
    Call Show(1, frmParent)
    
    If mblnChange Then
        ColConfig = mColConfigNew
    Else
        ColConfig = mColConfigOld
    End If
    
    ShowVsfColsListWindow = mblnChange
End Function



Private Sub CmdOK_Click()
'ȷ���������Բ���
On Error GoTo errHandle

    Dim i As Integer
    Dim strName As String
    Dim j As Integer
    Dim strCol() As String 'ÿһ������Ϣ
    
    Dim intCount As Integer
    Dim blnUserHide As Boolean

    intCount = UBound(mColConfigNew)
        
    For j = 0 To lstVsfColsName.ListCount - 1
        strName = lstVsfColsName.list(j)
        blnUserHide = IIf(lstVsfColsName.Selected(j), False, True)
        
        
        For i = 0 To intCount
            If strName = mColConfigNew(i).strName Then
                If mColConfigNew(i).blnIsUserHide <> blnUserHide Then mblnChange = True
                mColConfigNew(i).blnIsUserHide = blnUserHide
                Exit For
            End If
        Next
    Next
    
    Call Me.Hide
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Private Sub LoadColsList()
'����Ĭ����ʾ����
On Error GoTo errHandle
    Dim i As Integer
    Dim lngUbound As Long
    Dim strName As String
    Dim intTMP As Integer

    lngUbound = UBound(mColConfigNew)
    
    For i = 0 To lngUbound
        strName = mColConfigNew(i).strName
        
        If mColConfigNew(i).blnIsSysHide Then
            intTMP = intTMP + 1
        Else
            lstVsfColsName.list(i - intTMP) = strName
            lstVsfColsName.Selected(i - intTMP) = Not mColConfigNew(i).blnIsUserHide
        End If
    Next
    
    Exit Sub
errHandle:
    err.Raise -1, "��ͷ����", "LoadColsList" & vbCrLf & err.Description
'    Resume
End Sub


Private Sub cmdDefault_Click()
'�ָ�Ĭ�Ϲ�ѡ:���շ�������˳����ʾ����ȳ�ʼ��������״̬��ʼ��
On Error GoTo errHandle
    
    Call ReSetListHeadDefault
    Call LoadColsList
    mblnChange = True
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cmdExit_Click()
'ж�ش���
    mblnChange = False
    Unload Me
End Sub

Private Sub Form_Load()
    '�������ö�
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3

End Sub

Private Sub ReSetListHeadDefault()
'�ָ�Ĭ�ϣ�mColConfig �� �ֶ�˳��ָ���ϵͳĬ��  mColConfig���û�����ȫ��ȥ��
On Error GoTo errHandle
    Dim lngUbound As Integer
    Dim i As Integer
    Dim strName As String
    
    lngUbound = UBound(mColConfigNew)
    
    For i = 0 To lngUbound
        strName = mColConfigNew(i).strName
        mColConfigNew(i).lngColOrder = i + 1
        mColConfigNew(i).blnIsUserHide = False
        mColConfigNew(i).lngWidth = GetExtraWidth(strName) + 1.3 * mobjOwner.TextWidth(strName)
    Next
    Exit Sub
errHandle:
    err.Raise -1, "��ͷ����", "ReSetListHeadDefault" & vbCrLf & err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjOwner = Nothing
End Sub
