VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmCooperateUnitsReg 
   BackColor       =   &H8000000D&
   Caption         =   "��Լ��λ���ſ���"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   13245
   StartUpPosition =   1  '����������
   Begin VB.Frame fraLine 
      Height          =   88
      Left            =   1200
      TabIndex        =   17
      Top             =   2280
      Width           =   11535
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   120
      ScaleHeight     =   1530
      ScaleWidth      =   12720
      TabIndex        =   0
      Top             =   240
      Width           =   12720
      Begin VB.Frame fraInfo 
         Caption         =   "������Ϣ"
         Height          =   1260
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   12255
         Begin VB.CommandButton cmdOK 
            Caption         =   "ȷ��(&O)"
            Height          =   350
            Left            =   9360
            TabIndex        =   19
            Top             =   720
            Width           =   1100
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "ȡ��(&C)"
            Height          =   350
            Left            =   10575
            TabIndex        =   18
            Top             =   720
            Width           =   1100
         End
         Begin VB.TextBox txt�ű� 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   720
            MaxLength       =   5
            TabIndex        =   8
            Top             =   307
            Width           =   960
         End
         Begin VB.ComboBox cboItem 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3360
            TabIndex        =   7
            Text            =   "cboItem"
            Top             =   705
            Width           =   2235
         End
         Begin VB.ComboBox cboDoctor 
            Enabled         =   0   'False
            Height          =   300
            Left            =   6720
            TabIndex        =   6
            Top             =   705
            Width           =   2115
         End
         Begin VB.ComboBox cbo���� 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   720
            TabIndex        =   5
            Text            =   "cbo����"
            Top             =   705
            Width           =   2115
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�Һ�ʱ���뽨����"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5160
            TabIndex        =   4
            Top             =   360
            Width           =   1845
         End
         Begin VB.CheckBox chk��ſ��� 
            Caption         =   "��ſ���"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   3
            Top             =   330
            Width           =   1095
         End
         Begin VB.ComboBox cbo���� 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3360
            TabIndex        =   2
            Text            =   "cbo����"
            Top             =   307
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�ű�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   13
            Top             =   367
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   765
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��Ŀ"
            Height          =   180
            Left            =   3000
            TabIndex        =   11
            Top             =   765
            Width           =   360
         End
         Begin VB.Label lblҽ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ժ��ҽ��"
            Height          =   180
            Left            =   5940
            TabIndex        =   10
            Top             =   765
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   3000
            TabIndex        =   9
            Top             =   367
            Width           =   360
         End
      End
   End
   Begin VB.PictureBox picContent 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11160
      Left            =   5040
      ScaleHeight     =   11160
      ScaleWidth      =   8055
      TabIndex        =   15
      Top             =   2520
      Width           =   8055
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   16
         Top             =   -15
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   9870
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCooperateUnitsReg.frx":0000
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18283
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCooperateUnitsReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String

Private Const conPane_Info = 1
Private Const conPane_Plan = 3
Private mlngPriItem As Long
Private mlng����ID              As Long
Private mrs�޺�                 As ADODB.Recordset
Private mrs����                 As ADODB.Recordset
Private mstr�Ű�                As String '����|ȫ��||��һ|����||��������
Private mbln��ſ���            As Boolean
Private mblnUnload As Boolean
Private mblnʱ��                As Boolean '�������������ʱ�����ϸ���ʱ��������
Private mrsʱ���               As ADODB.Recordset
Private mstrKey       As String
Private mrsSource     As ADODB.Recordset
Private mrsUnitsReg   As ADODB.Recordset
Private WithEvents mfrmReg       As frmCooperateReg
Attribute mfrmReg.VB_VarHelpID = -1
Private WithEvents mfrmRegNoTime As frmCooperateRegArrange
Attribute mfrmRegNoTime.VB_VarHelpID = -1
Private mblnOk      As Boolean
 
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case conPane_Info
            Item.Handle = picList.hWnd
        Case conPane_Plan
            Item.Handle = picContent
    End Select
End Sub
 
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mbln��ſ��� Then
        If Not mfrmReg Is Nothing Then
            If mfrmReg.SaveData() = False Then Exit Sub
        End If
    Else
        If Not mfrmRegNoTime Is Nothing Then
             If mfrmRegNoTime.SaveData() = False Then Exit Sub
        End If
    End If
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Resize()
      Err.Number = 0
     On Error Resume Next
     With Me.picList
         .Left = Me.ScaleLeft
         .Top = Me.ScaleTop
         .Width = Me.ScaleWidth
     End With
     With fraLine
         .Left = Me.ScaleLeft
         .Top = picList.Top + picList.Height
         .Width = Me.ScaleWidth
     End With
     With Me.picContent
         .Left = Me.ScaleLeft
         .Top = picList.Top + picList.Height - 18 * Screen.TwipsPerPixelY
         .Height = ScaleHeight - .Top - Me.stbThis.Height
         .Width = ScaleWidth
     End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmReg Is Nothing Then Unload mfrmReg
    If Not mfrmRegNoTime Is Nothing Then Unload mfrmRegNoTime
    Set mfrmReg = Nothing
    Set mfrmRegNoTime = Nothing
     
End Sub

Public Function zlShowMe(ByVal lng����ID As Long, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���óɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-29 14:19:07
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mlng����ID = lng����ID
    If InitData() = False Then Exit Function
    InitPage
    Me.Show 1
    zlShowMe = mblnOk
End Function

Private Sub mfrmReg_frmUnload(ByVal blnCancel As Boolean)
    mblnOk = Not blnCancel
    Unload Me
End Sub

Private Sub mfrmRegNoTime_frmUnload(ByVal blnCancel As Boolean)
   mblnOk = Not blnCancel
   Unload Me
End Sub

Private Sub picContent_Resize()
    Err = 0: On Error Resume Next

    With picContent
        tbPage.Top = .ScaleTop
        tbPage.Left = .ScaleLeft
        'tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With

End Sub

Private Sub Form_Activate()
    Me.Icon = frmRegistPlan.Icon
    If mblnUnload Then mblnUnload = False: Unload Me
End Sub

'Private Sub InitPancel()
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��������
'    '����:
'    '����:2009-09-14 18:06:29
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim sngWidth As Single
'    Dim strReg   As String
'    Dim panThis  As Pane
'    Dim panInfo  As Pane
'
'    Set panInfo = dkpMan.CreatePane(conPane_Info, 900, 100, DockTopOf, Nothing)
'    'panThis.Title = "�ҺŰ�����Ϣ"
'    panInfo.Handle = picList.hWnd
'    panInfo.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'    panInfo.Tag = conPane_Info
'    panInfo.MaxTrackSize.Height = 100
'    panInfo.MinTrackSize.Height = 100
'    Set panThis = dkpMan.CreatePane(conPane_Plan, 160, 600, DockBottomOf, panInfo)
'    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
'    panThis.Title = "������λ"
'    panThis.Tag = conPane_Plan
'
'    panThis.Handle = picContent.hWnd
'    'dkpMan.Options.ThemedFloatingFrames = True
'    dkpMan.Options.ThemedFloatingFrames = False
'    dkpMan.Options.HideClient = False
'    dkpMan.Options.UseSplitterTracker = False 'ʵʱ�϶�
'    dkpMan.Options.AlphaDockingContext = True
'   ' panThis.MaxTrackSize.Height = 600
'     panThis.MinTrackSize.Height = 600
'
'    '    Set panThis = dkpMan.CreatePane(conPane_Plan, 250, 580, DockBottomOf, panThis)
'    '    panThis.Title = ""
'    '    panThis.Tag = conPane_Plan
'    '    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'    '    panThis.Handle = picPage.hWnd
'    '    dkpMan.Options.ThemedFloatingFrames = True
'    '    dkpMan.Options.HideClient = True
'    ' zlRestoreDockPanceToReg Me, dkpMan, "����"
'
'End Sub

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object

    Err = 0: On Error GoTo Errhand:
    
    tbPage.RemoveAll
    If mbln��ſ��� Then
    Set mfrmReg = New frmCooperateReg
        mfrmReg.frmInit mlng����ID
        Set ObjItem = tbPage.InsertItem(1, "", mfrmReg.hWnd, 0)
        ObjItem.Tag = 1
    Else
        Set mfrmRegNoTime = New frmCooperateRegArrange

        mfrmRegNoTime.frmInit mlng����ID, mlngModule, mstrPrivs
        Set ObjItem = tbPage.InsertItem(2, "", mfrmRegNoTime.hWnd, 0)
        ObjItem.Tag = 2
   End If
    With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        ' .PaintManager.Layout = xtpTabLayoutCompressed
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionTop
    End With
Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
 
'------------------------------------------------------------------------
'ҳ����ù����뷽��
'------------------------------------------------------------------------
Public Function InitData() As Boolean
    Dim strSQL As String
    Dim lng����ID       As Long
    Dim i       As Long
    Dim strTemp As String
    Dim rsTmp   As ADODB.Recordset
    If mlng����ID = -1 Then Exit Function
    lng����ID = mlng����ID
    On Error GoTo Hd
    strSQL = "Select count(0) as ��λ From �Һź�����λ  Where Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Val(Nvl(rsTmp!��λ)) = 0 Then MsgBox "û�����á��Һź�����λ��,�뵽�����ֵ�������!", vbOKOnly, Me.Caption: Exit Function
    strSQL = " " & _
    "   Select A.Id as ����ID,0 as �ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id," & _
    "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,nvl(A.Ĭ��ʱ�μ��,5) As Ĭ��ʱ�μ��, " & _
    "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ���� " & _
    "   From �ҺŰ��� A,�շ���ĿĿ¼ B,���ű� D " & _
    "   Where A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
    "         And A.Id=[1]"
    
    Set mrs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
         
    If mrs����.EOF Then
        ShowMsgbox "δ�ҵ�ָ���ĺű�,����!"
        Exit Function

    End If
        
    mstr�Ű� = ""

    For i = 0 To 6
        strTemp = Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")

        If Nvl(mrs����("��" & strTemp)) <> "" Then
            If mstr�Ű� <> "" Then mstr�Ű� = mstr�Ű� & "||"
            mstr�Ű� = mstr�Ű� & "��" & strTemp & "|" & Nvl(mrs����("��" & strTemp))
        End If
    Next

    If mstr�Ű� = "" Then Exit Function
'    strSQL = "Select ������Ŀ,�޺���,  ��Լ��,������Ŀ as ���� From  �ҺŰ������� where ����ID=[1]  Order BY ������Ŀ      "
'    Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    cbo����.Text = Nvl(mrs����!����)
    txt�ű�.Tag = Nvl(mrs����!����Id)
      
    txt�ű�.Text = Nvl(mrs����!����)
    cbo����.Text = Nvl(mrs����!����)
    cboItem.Text = Nvl(mrs����!��Ŀ)
    cboDoctor.Text = Nvl(mrs����!ҽ������)
    chk����.Value = IIf(Val(Nvl(mrs����!��������)) = 1, 1, 0)
    chk��ſ���.Value = IIf(Val(Nvl(mrs����!��ſ���)) = 1, 1, 0):  chk��ſ���.Tag = chk��ſ���.Value
    mbln��ſ��� = IIf(Val(Nvl(mrs����!��ſ���)) = 1, True, False)
    InitData = True
    Exit Function

Hd:

    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

