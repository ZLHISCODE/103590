VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmServiceEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   4590
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6885
   Icon            =   "frmServiceEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1470
      TabIndex        =   14
      Top             =   60
      Width           =   1575
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   4155
      Index           =   0
      Left            =   30
      ScaleHeight     =   4155
      ScaleWidth      =   7215
      TabIndex        =   12
      Top             =   510
      Width           =   7215
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4200
         TabIndex        =   10
         Top             =   3570
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(C)"
         Height          =   350
         Left            =   5535
         TabIndex        =   11
         Top             =   3570
         Width           =   1100
      End
      Begin VB.Frame fra 
         Height          =   3465
         Left            =   15
         TabIndex        =   13
         Top             =   -15
         Width           =   6840
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   4575
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2145
         End
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   5
            Left            =   615
            TabIndex        =   1
            Top             =   240
            Width           =   3420
         End
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   2
            Left            =   615
            TabIndex        =   5
            Top             =   630
            Width           =   3420
         End
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   3
            Left            =   4575
            TabIndex        =   7
            Top             =   630
            Width           =   2145
         End
         Begin VB.TextBox txt 
            Height          =   705
            Index           =   4
            Left            =   615
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   2670
            Width           =   6060
         End
         Begin VB.Frame Frame1 
            Height          =   1635
            Left            =   615
            TabIndex        =   16
            Top             =   960
            Width           =   6105
            Begin VB.OptionButton opt 
               Caption         =   "WebService"
               Height          =   225
               Index           =   1
               Left            =   1155
               TabIndex        =   18
               Top             =   30
               Width           =   1335
            End
            Begin VB.OptionButton opt 
               Caption         =   "Socket"
               Height          =   225
               Index           =   0
               Left            =   75
               TabIndex        =   17
               Top             =   30
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.PictureBox picBack 
               BorderStyle     =   0  'None
               Height          =   1275
               Index           =   0
               Left            =   30
               ScaleHeight     =   1275
               ScaleWidth      =   6045
               TabIndex        =   24
               Top             =   330
               Width           =   6045
               Begin VB.TextBox txt 
                  Height          =   345
                  Index           =   10
                  Left            =   3030
                  TabIndex        =   34
                  Top             =   885
                  Width           =   2895
               End
               Begin VB.TextBox txt 
                  Height          =   345
                  Index           =   9
                  Left            =   885
                  TabIndex        =   33
                  Top             =   900
                  Width           =   1245
               End
               Begin VB.TextBox txt 
                  Height          =   345
                  Index           =   8
                  Left            =   900
                  TabIndex        =   29
                  Top             =   495
                  Width           =   2070
               End
               Begin VB.TextBox txt 
                  Height          =   345
                  Index           =   7
                  Left            =   3825
                  TabIndex        =   28
                  Top             =   495
                  Width           =   2100
               End
               Begin VB.TextBox txt 
                  Height          =   315
                  Index           =   6
                  Left            =   885
                  TabIndex        =   25
                  Top             =   120
                  Width           =   5025
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "�ռ��ַ"
                  Height          =   180
                  Index           =   12
                  Left            =   2190
                  TabIndex        =   32
                  Top             =   960
                  Width           =   720
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "�ռ��ʶ"
                  Height          =   180
                  Index           =   11
                  Left            =   120
                  TabIndex        =   31
                  Top             =   960
                  Width           =   720
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "���÷���"
                  Height          =   180
                  Index           =   10
                  Left            =   105
                  TabIndex        =   30
                  Top             =   570
                  Width           =   720
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "��������"
                  Height          =   180
                  Index           =   9
                  Left            =   3030
                  TabIndex        =   27
                  Top             =   570
                  Width           =   720
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "���õ�ַ"
                  Height          =   180
                  Index           =   7
                  Left            =   105
                  TabIndex        =   26
                  Top             =   180
                  Width           =   720
               End
            End
            Begin VB.PictureBox picBack 
               BorderStyle     =   0  'None
               Height          =   855
               Index           =   1
               Left            =   90
               ScaleHeight     =   855
               ScaleWidth      =   5970
               TabIndex        =   19
               Top             =   390
               Width           =   5970
               Begin VB.TextBox txt 
                  Height          =   315
                  Index           =   1
                  Left            =   450
                  TabIndex        =   21
                  Top             =   510
                  Width           =   5310
               End
               Begin VB.TextBox txt 
                  Height          =   315
                  Index           =   0
                  Left            =   450
                  TabIndex        =   20
                  Top             =   90
                  Width           =   5310
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "PORT"
                  Height          =   180
                  Index           =   1
                  Left            =   30
                  TabIndex        =   23
                  Top             =   570
                  Width           =   360
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "IP"
                  Height          =   180
                  Index           =   0
                  Left            =   45
                  TabIndex        =   22
                  Top             =   150
                  Width           =   180
               End
            End
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�ӿ�"
            Height          =   180
            Index           =   8
            Left            =   165
            TabIndex        =   15
            Top             =   1050
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   6
            Left            =   4170
            TabIndex        =   2
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   5
            Left            =   165
            TabIndex        =   0
            Top             =   315
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   4
            Top             =   705
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�豸"
            Height          =   195
            Index           =   3
            Left            =   4170
            TabIndex        =   6
            Top             =   690
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "˵��"
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   8
            Top             =   2655
            Width           =   360
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmServiceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private mblnContiune As Boolean

Public Event AfterNewData(ByVal DataKey As String)
Public Event AfterModifyData(ByVal DataKey As String)
Public Event AfterDeleteData(ByVal DataKey As String)
Public Event Forward(ByRef DataKey As String, ByRef Cancel As Boolean)
Public Event Backward(ByRef DataKey As String, ByRef Cancel As Boolean)

'######################################################################################################################

Public Function InitDialog(ByVal frmParent As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    
    InitDialog = True
    
End Function

Public Sub NewData()
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 1
    Me.Caption = "������������"
    mstrDataKey = ""
    
    Call InitData
    Call InitCommandBar
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub ModifyData(ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    
    mbytMode = 2
    mstrDataKey = strDataKey
    
    Me.Caption = "�޸ķ�������"

    Call InitData
    Call InitCommandBar
    
    Call ReadData(mstrDataKey)
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub DeleteData(ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 3
    If strDataKey = "" Then Exit Sub
    mstrDataKey = strDataKey
    
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "ID", mstrDataKey)
        
    If gclsMsgBase.EditService("Delete", mrsPara) Then
        RaiseEvent AfterDeleteData(mstrDataKey)
    End If
End Sub

'######################################################################################################################
Private Function InitData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    mblnContiune = False
    
    Set rsTmp = gclsMsgBase.GetServiceStruct()
    If Not (rsTmp Is Nothing) Then
        txt(0).MaxLength = 15
        txt(1).MaxLength = 5
        txt(2).MaxLength = rsTmp("app").DefinedSize
        txt(3).MaxLength = rsTmp("device").DefinedSize
        txt(4).MaxLength = rsTmp("note").DefinedSize
        txt(5).MaxLength = rsTmp("title").DefinedSize
        txt(6).MaxLength = rsTmp("interface_para").DefinedSize
        txt(7).MaxLength = 50
        txt(8).MaxLength = 100
        txt(9).MaxLength = 20
        txt(10).MaxLength = 50
    End If
    
    With cbo(0)
        .AddItem "1-Ŀ�����"
        .ItemData(.NewIndex) = 1
        
        .AddItem "2-���շ���"
        .ItemData(.NewIndex) = 2
        
        .ListIndex = 0
    End With
    
    picBack(0).Visible = False
    picBack(1).Visible = True
        
    InitData = True
End Function

Private Function ReadData(ByVal strDataKey As String) As Boolean

    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strTemp As String
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim varTemp As Variant
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)
    
    mblnReading = True
    Set rsTmp = gclsMsgBase.GetService("id", rsCondition)
    If rsTmp.BOF = False Then

        txt(2).Text = zlCommFun.NVL(rsTmp("app").Value)
        txt(3).Text = zlCommFun.NVL(rsTmp("device").Value)
        txt(4).Text = zlCommFun.NVL(rsTmp("note").Value)
        txt(5).Text = zlCommFun.NVL(rsTmp("title").Value)
        strTemp = zlCommFun.NVL(rsTmp("interface_para").Value)
        Select Case zlCommFun.NVL(rsTmp("interface_type").Value)
        Case 1
            opt(0).Value = True
            If InStr(strTemp, "/") > 0 Then
                txt(0).Text = Mid(strTemp, 1, InStr(strTemp, "/") - 1)
                txt(1).Text = Mid(strTemp, InStr(strTemp, "/") + 1)
            End If
        Case 2
            '<root><address></address><method></method></root>
            
            varTemp = Split(strTemp, vbCrLf)
    
            txt(6).Text = varTemp(0)
            If UBound(varTemp) > 0 Then txt(8).Text = varTemp(1)
            If UBound(varTemp) > 1 Then txt(7).Text = varTemp(2)
            If UBound(varTemp) > 2 Then txt(9).Text = varTemp(3)
            If UBound(varTemp) > 3 Then txt(10).Text = varTemp(4)
            
            opt(1).Value = True
            
        End Select
        
        Call zlControl.CboLocate(cbo(0), zlCommFun.NVL(rsTmp("kind").Value), True)
        
    End If
    mblnReading = False
    mblnDataChanged = False
    
    ReadData = True
    
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    

    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Select Case mbytMode
    Case 1
        Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, "ȷ�����������", True)
    Case 2
        Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, "ȷ��������޸�", True)
    End Select
    objControl.IconId = conMenu_View_UnCheck
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, conMenu_View_LocationItem, "����", True, , xtpButtonIconAndCaption)
    objControl.IconId = conMenu_View_Find

    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "����")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Forward, "��һ��")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Backward, "��һ��")


    
    txtLocation.Visible = (mbytMode = 2)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    
    If Len(Trim(txt(5).Text)) = 0 Then
        ShowSimpleMsg "��������Ʋ���Ϊ�գ�"
        Call LocationObj(txt(5))
        Exit Function
    End If
    
    If zlCommFun.CheckValid(txt(0).Text, emDataFormat.IP4) = False Then
        ShowSimpleMsg "�����ַ��Ч�������ַ��ʽΪ��[0-255].[0-255].[0-255].[0-255]����192.168.4.51��"
        Call LocationObj(txt(0))
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Function SaveData(ByRef strDataKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
    
    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
    
    If opt(0).Value Then
        Call zlCommFun.SetParameter(rsPara, "interface_type", 1)
        Call zlCommFun.SetParameter(rsPara, "interface_para", Trim(txt(0).Text) & "/" & Val(txt(1).Text))
    Else
        Call zlCommFun.SetParameter(rsPara, "interface_type", 2)
        Call zlCommFun.SetParameter(rsPara, "interface_para", Trim(txt(6).Text) & vbCrLf & Trim(txt(8).Text) & vbCrLf & Trim(txt(7).Text) & vbCrLf & Trim(txt(9).Text) & vbCrLf & Trim(txt(10).Text))
    End If
    Call zlCommFun.SetParameter(rsPara, "app", Trim(txt(2).Text))
    Call zlCommFun.SetParameter(rsPara, "device", Trim(txt(3).Text))
    Call zlCommFun.SetParameter(rsPara, "note", Trim(txt(4).Text))
    Call zlCommFun.SetParameter(rsPara, "title", Trim(txt(5).Text))
    Call zlCommFun.SetParameter(rsPara, "kind", Trim(cbo(0).ItemData(cbo(0).ListIndex)))
        
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '����
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        
        SaveData = gclsMsgBase.EditService("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�޸�
        SaveData = gclsMsgBase.EditService("UPDATE", rsPara)
    End Select
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Click(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Dim blnCancel As Boolean
    Dim strDataKey As String
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward               '��һ��
        
        strDataKey = mstrDataKey
        
        RaiseEvent Forward(strDataKey, blnCancel)
        If blnCancel = False Then
        
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
    
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward               '��һ��
        
        strDataKey = mstrDataKey
        
        RaiseEvent Backward(strDataKey, blnCancel)
        If blnCancel = False Then
            
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        
        Dim strText As String
        Dim rsCondition As ADODB.Recordset
        Dim rsData As ADODB.Recordset
        Dim rs As ADODB.Recordset
        
        If txtLocation.Text <> "" Then
            
            txtLocation.Tag = ""
            
            strText = UCase(txtLocation.Text)
                        
            Set rsCondition = zlCommFun.CreateCondition
            
            Call zlCommFun.SetCondition(rsCondition, "server_ip", "%" & strText & "%")
            Set rsData = gclsMsgBase.GetService("title", rsCondition)
            
            If zlCommFun.ShowPubSelect(Me, txtLocation, 2, "����,1200,0,1;����,2400,0,1;��ַ,1500,0,1;�˿�,600,0,0;����,900,0,0;�豸,900,0,0", Me.Name & "\�������ù���", "����±���ѡ��һ����������", rsData, rs, , , , , , True) = 1 Then
                mstrDataKey = rs("id").Value
                Call ReadData(mstrDataKey)
                txtLocation.Tag = ""
            Else
                txtLocation.Tag = ""
                Call LocationObj(txtLocation, True)
                Exit Sub
            End If
                        
            Call LocationObj(txtLocation, True)
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        mblnContiune = Not mblnContiune
    End Select
    
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '���������ؼ�Resize����
    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case conMenu_View_Filter, conMenu_View_LocationItem, conMenu_View_Backward, conMenu_View_Forward, 0
        Control.Visible = (mbytMode = 2)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnContiune
        Control.IconId = IIf(mblnContiune = True, conMenu_View_Check, conMenu_View_UnCheck)
    End Select
    
End Sub

Private Sub chk_Click(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    '
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strOldDataKey As String
    Dim rsTmp As ADODB.Recordset
    
    If mblnDataChanged = True Then
        If ValidData = True Then
    
            If SaveData(mstrDataKey) = True Then
                
                If strOldDataKey <> "" Then
                    RaiseEvent AfterModifyData(strOldDataKey)
                End If
                
                Select Case mbytMode
                Case 1
                    RaiseEvent AfterNewData(mstrDataKey)
                Case 2
                    RaiseEvent AfterModifyData(mstrDataKey)
                End Select
                
                If mblnContiune = False Then
                    mblnDataChanged = False
                    Unload Me
                Else
                    '���û�����������һ������״̬
                    If mbytMode = 1 Then
                        mstrDataKey = ""
                        txt(0).Text = ""
                        txt(1).Text = ""
                        txt(2).Text = ""
                        txt(3).Text = ""
                        txt(4).Text = ""
                        txt(5).Text = ""
                        txt(6).Text = ""
                        txt(7).Text = ""
                        txt(8).Text = ""
                        txt(9).Text = ""
                        txt(10).Text = ""
                    End If
                    Call LocationObj(txt(0))
                    mblnDataChanged = False
                End If
                
            End If
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbNo)
        If Cancel Then Exit Sub
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    
    Select Case Index
    Case 0
        picBack(0).Visible = False
        picBack(1).Visible = True
    Case 1
        picBack(0).Visible = True
        picBack(1).Visible = False
    End Select
    
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    mblnDataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 4
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
        '
        
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim bytMode As Byte
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 4
        zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub txtLocation_Change()
    txtLocation.Tag = "Changed"
End Sub

Private Sub txtLocation_GotFocus()
    zlControl.TxtSelAll txtLocation
End Sub

Private Sub txtLocation_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        KeyCode = 0
        txtLocation.Text = ""
        txtLocation.Tag = ""
    End If

End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        If txtLocation.Text <> "" Then
            txtLocation.Tag = ""
            
            Dim obj As CommandBarControl
            
            Set obj = cbsMain.FindControl(, conMenu_View_Filter, True)
            If obj.Enabled = True Then
                Call cbsMain_Execute(obj)
            End If

        End If
        txtLocation.Tag = ""
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txtLocation_Validate(Cancel As Boolean)
    If (txtLocation.Tag = "Changed") Then
        txtLocation.Tag = ""
    End If
End Sub
