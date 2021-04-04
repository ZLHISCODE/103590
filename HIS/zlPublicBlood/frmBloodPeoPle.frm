VERSION 5.00
Begin VB.Form frmBloodPeoPle 
   BorderStyle     =   0  'None
   Caption         =   "frmBloodPeoPle"
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin zlPublicBlood.usrCardPeople UCP 
      Height          =   2550
      Left            =   525
      TabIndex        =   0
      Top             =   660
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   4498
   End
End
Attribute VB_Name = "frmBloodPeoPle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CardChanged()
Public Event AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String, ByVal blnNext As Boolean, blnfind As Boolean) '���ҵ�IDKindStr���濨Ƭ�ϣ��򷵻��¼��е���������
Public Event CodeFilter(ByVal strCode As String)
Public strReturn As String
Private m_CanCheck As Boolean
Private m_FindStart As Boolean

Public Sub ShowPeople(Optional ByVal rsBR As ADODB.Recordset)
    '���ܣ����øÿؼ��ķ������ܹ�δ�ؼ��ṩ��ʼ�Ĺ���������
    '������rsBRҪ��ʾ������Դ������Դ��Ҫ����ID������ֵ�л᷵��ID�ţ�����id��Ϊ�˷����û���ѯ��
    UCP.ShowPeople rsBR
End Sub

Public Sub UserInit(ByVal frmMain As Object, str���� As String, Optional ByVal imgList As Object, Optional ByVal lngModule As Long = 0, Optional ByVal strIDKindstr As String = "")
    '���������Ҫ��һ���ַ������������ɫ���ݣ���ð���ɫд�ڵ�һ������Ϊ�����λ���Ǻ�ҳ��ؼ�λ�ö�Ӧ��
    UCP.UserInit frmMain, str����, imgList, lngModule, strIDKindstr
End Sub

Public Sub Hook()
    Set gobjFScrollBar = UCP.FScrollBar
    glngBooldPepWinProc = GetWindowLong(UCP.objPicBack.hWnd, GWL_WNDPROC)
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Public Sub UnHook()
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, glngBooldPepWinProc
End Sub

Private Sub Form_Resize()
    '���ܣ����ƿؼ���λ��
    UCP.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Function GetCheckedData() As ADODB.Recordset
    '���ܣ���ȡ���ѡ�������
    Set GetCheckedData = UCP.GetCheckedData
End Function

Private Sub UCP_AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String, ByVal blnNext As Boolean, blnfind As Boolean)
    RaiseEvent AfterPatiFind(strIDKindstr, strValue, blnNext, blnfind)
End Sub

Private Sub UCP_CodeFilter(ByVal strCode As String)
    RaiseEvent CodeFilter(strCode)
End Sub

Public Sub CodeFilter(rs As Recordset)
    UCP.FilterRefreshByCode rs
End Sub

Private Sub UCP_CardChanged()
    '���ܣ���ȡѡ��ѡ�������
    strReturn = UCP.strReturn
    RaiseEvent CardChanged
End Sub

'��ȡcancheck����
Public Property Get CanCheck() As Boolean
    CanCheck = m_CanCheck
    UCP.CanCheck = m_CanCheck
End Property
Public Property Let CanCheck(newCanCheck As Boolean)
    m_CanCheck = newCanCheck
    UCP.CanCheck = m_CanCheck
End Property

Public Property Let FindStart(newFindStart As Boolean)
    '���ܣ���ʼ����ѯ
    m_FindStart = newFindStart
    UCP.FindStart = m_FindStart
End Property

Public Sub FindPatiByVbKey(Optional ByVal blnNext As Boolean)
    '���ܣ�ͨ����ݼ��ķ�ʽ��ʼ���ң��������һ������
    Call UCP.FindPatiByVbKey(blnNext)
End Sub

Public Sub SetPIFocus()
    '���ܣ���λ����ѯ�ؼ�
    Call UCP.SetPIFocus
End Sub

Public Sub SetCardFocus(strTitle As String, strfind As String)
    '��λ��ָ������Ա����
    Call UCP.SetCardFocus(strTitle, strfind)
End Sub

Public Sub FilterRefreshBC(rs As Recordset)
    Call UCP.FilterRefreshByCode(rs)
End Sub
