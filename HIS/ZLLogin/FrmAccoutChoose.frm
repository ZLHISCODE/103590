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
Private mrsSystems As ADODB.Recordset
Private mstrSQL As String
Private mstrCodes As String
Private mstrComponent As String
Private mlngCur As Long
Private mintCurTab As Integer
Private mblnMutil As Boolean
Private mblnMutilSys As Boolean

Public BlnSelect As Boolean

Private Sub Cmdȡ��_Click()
    gclsLogin.IsCancel = True
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    '����SQL
    Dim lvwThis As Control, LvwItem As ListItem
    On Error GoTo ErrH
    For Each lvwThis In Me.Controls
        If TypeName(lvwThis) = "ListView" Then
            If lvwThis.Index <> 0 Then
                mstrSQL = mstrSQL & IIf(mstrSQL = "", "", ",") & "'" & lvwThis.SelectedItem.Tag & "'"
            Else
                For Each LvwItem In lvwThis.ListItems
                    mstrSQL = mstrSQL & IIf(mstrSQL = "", "", ",") & "'" & LvwItem.Tag & "'"
                Next
            End If
        End If
    Next
    
    '���û���κ�ϵͳ��ѡ�������Ƿ���ڱ����ִ��
    If mstrSQL = "" Then
        mstrSQL = "Select 1" & vbNewLine & _
                "From zlProgFuncs" & vbNewLine & _
                "Where ϵͳ Is Null And ��� In (Select Distinct ���" & vbNewLine & _
                "                            From zlRoleGrant G, zlUserRoles S" & vbNewLine & _
                "                            Where g.��ɫ = s.��ɫ And s.�û� = [1] And ϵͳ Is Null) And Rownum < 2"
        Set mrsSystems = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���ñ���", gclsLogin.DBUser)
        mstrSQL = ""
        If Not mrsSystems.EOF Then mstrSQL = "REPORT"
    End If
    
    BlnSelect = True
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    If BlnSelect = False Then
        Dim LngStyle As Long
        LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
        LngStyle = LngStyle Or WinStyle
        Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
        
        ShowWindow Me.hwnd, 0 '������
        ShowWindow Me.hwnd, 1 '����ʾ
    End If
End Sub

Private Sub Form_Load()
    Dim blnMutilAccout As Boolean
    
    On Error GoTo ErrH
    Me.Hide
    mblnMutilSys = False
    BlnSelect = False
    blnMutilAccout = False
    mstrComponent = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
        
    mstrSQL = "Select 1 From zlSystems Where ������ = [1]"
    Set mrsSystems = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ����ϵͳ", gclsLogin.DBUser)
    gclsLogin.IsSysOwner = mrsSystems.RecordCount > 0
    mstrSQL = "Select 1 From Zlsystems Where Mod(���, 100) <> 0"
    Set mrsSystems = zlDatabase.OpenSQLRecord(mstrSQL, "�Ƿ���ڷǱ�׼����", gclsLogin.DBUser)
    blnMutilAccout = mrsSystems.RecordCount > 0
    
    If gclsLogin.IsSysOwner Then
        '������ֻ��������Լ���ϵͳ(��Ϊ�ж������������)
        '����ȡ��������߼����ϲ�������֧����Ҫ������һ��ϵͳ������������һ����ϵͳ����Ȩ�û���
        mstrSQL = " Select Distinct g.ϵͳ From zlRoleGrant G, zlUserRoles S Where g.��ɫ = s.��ɫ and s.�û� = [1] And g.���� = '����'  Union Select ��� From zlSystems Where ������ = [1] "
    Else
        '��ͨ�û�ֻ�������������ɫ��Ȩ���������ϵͳ
        mstrSQL = " Select Distinct g.ϵͳ From zlRoleGrant G, zlUserRoles S Where g.��ɫ = s.��ɫ and s.�û� = [1]  And g.���� = '����'"
    End If
    mstrSQL = "Select Substr(LPad(���, 5, '0'), 4) ���, ��� ϵͳ, ����" & vbNewLine & _
                    "From zlSystems" & vbNewLine & _
                    "Where ��� In" & vbNewLine & _
                    "      (Select Distinct p.ϵͳ" & vbNewLine & _
                    "       From zlPrograms P," & vbNewLine & _
                    "      (" & mstrSQL & ") f" & vbNewLine & _
                    "       WHERE p.ϵͳ = f.ϵͳ " & vbNewLine & _
                    IIf(mstrComponent <> "", "    AND  Upper(p.����) IN (" & mstrComponent & ")) " & vbNewLine, ")") & _
                    " ORDER BY ����, ���"
                    
    '�򿪼�¼��������޶����ף����˳�
    Set mrsSystems = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ����ϵͳ", gclsLogin.DBUser)
    With mrsSystems
        mintCurTab = 0
        mstrCodes = ""
        
        Do While Not .EOF
            '����ϵͳ�Ƿ��ж�����,�������Index=0��Listview;��������Listview,����
            mblnMutil = False
            mlngCur = .AbsolutePosition
            If mstrCodes <> !���� Then
                mstrCodes = !����
                .Filter = "����='" & mstrCodes & "'"
                mblnMutil = (.RecordCount > 1)
                If mblnMutilSys = False Then mblnMutilSys = mblnMutil
                
                If mblnMutil Then
                    mintCurTab = mintCurTab + 1
                    Load LvwSelect(mintCurTab)
                    With LvwSelect(mintCurTab)
                        .ListItems.Clear
                        .Left = LvwSelect(mintCurTab - 1).Left
                        .Top = LvwSelect(mintCurTab - 1).Top + 1400
                        .Width = LvwSelect(mintCurTab - 1).Width
                        .Height = LvwSelect(mintCurTab - 1).Height
                        .Visible = True
                    End With
                    Load LblNote(mintCurTab)
                    With LblNote(mintCurTab)
                        .Left = LblNote(mintCurTab - 1).Left
                        .Top = LblNote(mintCurTab - 1).Top + 1400
                        .Width = LblNote(mintCurTab - 1).Width
                        .Height = LblNote(mintCurTab - 1).Height
                        .Visible = True
                        .Caption = mstrCodes
                    End With
                    
                    '�����¼
                    Do While Not .EOF
                        LvwSelect(mintCurTab).ListItems.Add , "K_" & LvwSelect(mintCurTab).ListItems.Count + 1, mstrCodes & IIf(Val(!���) = 0, "", "��" & Val(!���) & "��"), 1
                        LvwSelect(mintCurTab).ListItems("K_" & LvwSelect(mintCurTab).ListItems.Count).Tag = !ϵͳ
                        .MoveNext
                    Loop
                Else
                    '�����¼��LvwSelect(0)
                    LvwSelect(0).ListItems.Add , "K_" & LvwSelect(0).ListItems.Count + 1, mstrCodes & IIf(Val(!���) = 0, "", "��" & Val(!���) & "��"), 1
                    LvwSelect(0).ListItems("K_" & LvwSelect(0).ListItems.Count).Tag = !ϵͳ
                End If
            End If
                
            .Filter = 0
            .MoveFirst
            .Move mlngCur - 1
            .MoveNext
        Loop
        
        With Cmdȷ��
            .Top = LvwSelect(mintCurTab).Top + LvwSelect(mintCurTab).Height + 150
        End With
        Cmdȡ��.Top = Cmdȷ��.Top
        
        Me.Height = Me.Cmdȷ��.Top + Me.Cmdȷ��.Height + 550
    End With
    
    mstrSQL = ""
    If mblnMutilSys = False Then Cmdȷ��_Click
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function Show_me() As String
    On Error Resume Next
    
    Me.Show 1
    Show_me = mstrSQL
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mrsSystems = Nothing
End Sub

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


