VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAccoutChoose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "FrmAccoutChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ListView LvwSelect 
      Height          =   1005
      Index           =   0
      Left            =   1380
      TabIndex        =   2
      Top             =   -600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1773
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "Img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Img 
      Left            =   4860
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccoutChoose.frx":062A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   1
      Top             =   1860
      Width           =   1100
   End
   Begin VB.CommandButton Cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2910
      TabIndex        =   0
      Top             =   1860
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   240
      Picture         =   "FrmAccoutChoose.frx":0C64
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "    ������ͬʱ��������ϵͳ������׵�Ȩ�ޣ���ѡ�񱾴β��������ף�"
      Height          =   405
      Left            =   990
      TabIndex        =   4
      Top             =   60
      Width           =   4455
   End
   Begin VB.Label LblNote 
      AutoSize        =   -1  'True
      Caption         =   "ҽԺ��Ϣϵͳ"
      Height          =   180
      Index           =   0
      Left            =   1350
      TabIndex        =   3
      Top             =   -780
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "FrmAccoutChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RecSystems As New ADODB.Recordset
Private strSQL As String
Private strCodes As String
Private StrComponent As String
Private LngCur As Long
Private IntCurTab As Integer
Private BlnMutil As Boolean
Private BlnMutilSys As Boolean
Public BlnSelect As Boolean

Private Sub Cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    '����SQL
    Dim lvwThis As Control, LvwItem As ListItem
    
    For Each lvwThis In Me.Controls
        If TypeName(lvwThis) = "ListView" Then
            If lvwThis.Index <> 0 Then
                strSQL = strSQL & IIf(strSQL = "", "", ",") & "'" & lvwThis.SelectedItem.Tag & "'"
            Else
                For Each LvwItem In lvwThis.ListItems
                    strSQL = strSQL & IIf(strSQL = "", "", ",") & "'" & LvwItem.Tag & "'"
                Next
            End If
        End If
    Next
    
    '���û���κ�ϵͳ��ѡ�������Ƿ���ڱ����ִ��
    If strSQL = "" Then
        
        Set RecSystems = gclsMsgSystem.GetProgFuncsCount
        With RecSystems
'            If .State = 1 Then .Close
'            strSQL = "   Select Count(*) Records From zlprogfuncs " & _
'                     "   Where ϵͳ Is Null" & _
'                     "   And ��� in (Select ��� From zlRoleGrant G,session_roles S Where G.��ɫ=S.Role)"
'            .Open strSQL, gcnOracle
            strSQL = ""
            
            If Not .EOF Then
                If !Records <> 0 Then
                    strSQL = "REPORT"
                End If
            End If
        End With
    End If
    
    BlnSelect = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If BlnSelect = False Then
        Dim LngStyle As Long
        LngStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
        LngStyle = LngStyle Or WinStyle
        Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, LngStyle)
        
        ShowWindow Me.hWnd, 0 '������
        ShowWindow Me.hWnd, 1 '����ʾ
    End If
End Sub

Private Sub Form_Load()
    Me.Hide
    BlnMutilSys = False
    BlnSelect = False
    
'    strCodes = zlRegFunctions(gstrRegCode) & " Or ϵͳ Is NULL "
    StrComponent = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
        
    '�򿪼�¼��������޶����ף����˳�
    
    Set RecSystems = gclsMsgSystem.GetAccount(StrComponent)
    With RecSystems
'        If .State = 1 Then .Close
'        .Open strSQL, gcnOracle
        IntCurTab = 0
        strCodes = ""
        
        Do While Not .EOF
            '����ϵͳ�Ƿ��ж�����,�������Index=0��Listview;��������Listview,������
            BlnMutil = False
            LngCur = .AbsolutePosition
            If strCodes <> !���� Then
                strCodes = !����
                .Filter = "����='" & strCodes & "'"
                BlnMutil = (.RecordCount > 1)
                If BlnMutilSys = False Then BlnMutilSys = BlnMutil
                
                If BlnMutil Then
                    IntCurTab = IntCurTab + 1
                    Load LvwSelect(IntCurTab)
                    With LvwSelect(IntCurTab)
                        .ListItems.Clear
                        .Left = LvwSelect(IntCurTab - 1).Left
                        .Top = LvwSelect(IntCurTab - 1).Top + 1400
                        .Width = LvwSelect(IntCurTab - 1).Width
                        .Height = LvwSelect(IntCurTab - 1).Height
                        .Visible = True
                    End With
                    Load LblNote(IntCurTab)
                    With LblNote(IntCurTab)
                        .Left = LblNote(IntCurTab - 1).Left
                        .Top = LblNote(IntCurTab - 1).Top + 1400
                        .Width = LblNote(IntCurTab - 1).Width
                        .Height = LblNote(IntCurTab - 1).Height
                        .Visible = True
                        .Caption = strCodes
                    End With
                    
                    '�����¼
                    Do While Not .EOF
                        LvwSelect(IntCurTab).ListItems.Add , "K_" & LvwSelect(IntCurTab).ListItems.Count + 1, strCodes & IIf(Val(!���) = 0, "", "��" & Val(!���) & "��"), 1
                        LvwSelect(IntCurTab).ListItems("K_" & LvwSelect(IntCurTab).ListItems.Count).Tag = !ϵͳ
                        .MoveNext
                    Loop
                Else
                    '�����¼��LvwSelect(0)
                    LvwSelect(0).ListItems.Add , "K_" & LvwSelect(0).ListItems.Count + 1, strCodes & IIf(Val(!���) = 0, "", "��" & Val(!���) & "��"), 1
                    LvwSelect(0).ListItems("K_" & LvwSelect(0).ListItems.Count).Tag = !ϵͳ
                End If
            End If
                
            .Filter = 0
            .MoveFirst
            .Move LngCur - 1
            .MoveNext
        Loop
        
        With Cmdȷ��
            .Top = LvwSelect(IntCurTab).Top + LvwSelect(IntCurTab).Height + 150
        End With
        Cmdȡ��.Top = Cmdȷ��.Top
        
        Me.Height = Me.Cmdȷ��.Top + Me.Cmdȷ��.Height + 550
    End With
    
    strSQL = ""
    If BlnMutilSys = False Then Cmdȷ��_Click
End Sub

Public Function Show_me() As String
    On Error Resume Next
    
    Me.Show 1
    Show_me = strSQL
End Function

Private Sub LvwSelect_DblClick(Index As Integer)
    LvwSelect_KeyDown Index, vbKeyReturn, 0
End Sub

Private Sub LvwSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < LvwSelect.Count - 1 Then
            LvwSelect(Index + 1).SetFocus
        Else
            Cmdȷ��.SetFocus
        End If
    End If
End Sub
