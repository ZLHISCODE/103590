VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiCureCardTypeMgr 
   Caption         =   "ҽ�ƿ�������"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9930
   Icon            =   "frmPatiCureCardTypeMgr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   9930
   StartUpPosition =   1  '����������
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiCureCardTypeMgr.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12435
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
   Begin MSComctlLib.ImageList ils32 
      Left            =   1350
      Top             =   1890
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardTypeMgr.frx":115E
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardTypeMgr.frx":1A38
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1440
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardTypeMgr.frx":2312
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardTypeMgr.frx":28AC
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2235
      Left            =   1950
      TabIndex        =   1
      Top             =   1440
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   225
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmPatiCureCardTypeMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mlngModule As Long, mstrPrivs As String
Private Const mstrLvw As String = "����,2174.74,0,1;����,799.9371,0,2;����,629.8583,2,0;ǰ׺�ı�,929.7639,2,0;" & _
    "���ų���,989.8583,0,1;ȱʡ,630,2,1;�̶���,760,2,0;�ϸ����,1000,2,0;�����,800,2,0;�����ʻ�,1000,2,0;" & _
    "ȫ��,1400,2,0;����,1500,0,0;ҽ�ƿ���,1500,0,0;���㷽ʽ,1620,0,0;��������,1000,0,0;����,600,2,0;��ע,2000,0,0;" & _
    "ģ������,1000,2,0;�Ƿ��ƿ�,1000,2,0;�Ƿ񷢿�,1000,2,0;�Ƿ�д��,1000,2,0;ת�ʼ�����,1200,2,0;ˢ��,800,2,0;" & _
    "ɨ�迨,800,2,0;�Ӵ�ʽ����,1200,2,0;�ǽӴ�����,1200,2,0;��������,1000,2,0;����ֿ�����,1300,2,0;" & _
    "���͵��ýӿ�,1300,2,0;�Ƿ��˿��鿨,1300,2,0;֤��,800,2,0;���ûس�,1000,2,0;�Ƿ�ȱʡ����,1000,2,0" '�����:56508
Private mintColumn As Integer
Private mblnShowStop As Boolean

Private Sub LoadData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2011-06-27 20:52:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objItem As ListItem, lngCol As Long, varValue As Variant
    Dim strͣ�� As String, strKey As String
    
    On Error GoTo errHandle
    
    '"           nvl(���볤��,10) as ���볤��,nvl(���볤������,0) as ���볤������,nvl(�������,0) as �������," & _
    '"����,2000,0,1;����,800,0,2;����,500,0,0;ǰ׺�ı�,800,0,0;���ų���,800,0,1;ȱʡ��־,400,0,1;�Ƿ�̶�,400,0,0;�Ƿ��ϸ����,1000,0,0;�Ƿ�ˢ��,1000,0,0;�Ƿ�����,1000,0,0;�Ƿ�����ʻ�,1400,0,0;�Ƿ�ȫ��,1400,0,0;����,1500,0,0;�ض���Ŀ,1000,0,0;���㷽ʽ,1000,0,0;��������,1000,0,0;�Ƿ�����,1000,0,0;��ע,2000,0,0"
    '�����:56508
    '103310,���ϴ�,2016/12/7:���ź����ӻس���λ
    '90875:���ϴ�,2016/1/21,����ҽ�ƿ�֤������
    '77872,���ϴ�,2014/9/15:�Ƿ�֧��ת�ʼ�����
    strSQL = "" & _
    "   Select A.Id, A.����, A.����, A.����, A.ǰ׺�ı�, A.���ų���, decode(nvl(A.ȱʡ��־,0),1,'��','') as ȱʡ,  " & _
    "           decode(nvl(A.�Ƿ�̶�,0),1,'��','') as �̶���, decode(nvl(A.�Ƿ��ϸ����,0),1,'��','') as  �ϸ����, " & _
    "           decode(nvl(A.�Ƿ�����,0),1,'Ժ�ڿ�','Ժ�⿨') as    �����," & _
    "           decode(nvl(A.�Ƿ�����ʻ�,0),1,'��','') as     �����ʻ�, decode(nvl(A.�Ƿ�ȫ��,0),1,'��','') as    ȫ��," & _
    "           A.����,C.���� as ҽ�ƿ���, A.���㷽ʽ,A.��������, decode(nvl(A.�Ƿ�����,0),1,'��','') as   ����, A.��ע,  " & _
    "           decode(nvl(A.�Ƿ�ģ������,0),1,'��','')  as ģ������," & _
    "           decode(nvl(A.�Ƿ��ƿ�,0),1,'��','') as   �Ƿ��ƿ�,decode(nvl(A.�Ƿ񷢿�,0),1,'��','') as   �Ƿ񷢿�,decode(nvl(A.�Ƿ�д��,0),1,'��','') as   �Ƿ�д��," & _
    "           decode(nvl(A.�Ƿ�ת�ʼ�����,0),1,'��','') as   ת�ʼ�����, decode(nvl(A.�Ƿ�֤��,0),1,'��','') as   ֤��," & _
    "           decode(substr(nvl(A.��������,'0000'),1,1),1,'��','') as   ˢ��," & _
    "           decode(substr(nvl(A.��������,'0000'),2,1),1,'��','') as   ɨ�迨," & _
    "           decode(substr(nvl(A.��������,'0000'),3,1),1,'��','') as   �Ӵ�ʽ����," & _
    "           decode(substr(nvl(A.��������,'0000'),4,1),1,'��','') as   �ǽӴ�����," & _
    "           decode(nvl(A.���̿��Ʒ�ʽ,0),0,'����',1,'����',2,'�ַ�','����') as  ��������, " & _
    "           decode(nvl(A.�Ƿ�ֿ�����,0),1,'��','') as ����ֿ�����, " & _
    "           decode(nvl(A.���͵��ýӿ�,0),1,'��','') as ���͵��ýӿ�, " & _
    "           Decode(Nvl(a.�Ƿ��˿��鿨,0),1,'��','') As �Ƿ��˿��鿨, " & _
    "           decode(A.�豸�Ƿ����ûس�,1,'��','') as   ���ûس�, " & _
    "           decode(A.�Ƿ�ȱʡ����,1,'��','') as   �Ƿ�ȱʡ���� " & _
    "    From ҽ�ƿ���� A ,�շ��ض���Ŀ B,�շ���ĿĿ¼ C" & _
    "    Where   A.�ض���Ŀ=B.�ض���Ŀ(+) and B.�շ�ϸĿID=C.ID(+) " & _
            IIf(mblnShowStop, "", " and Nvl(�Ƿ�����,0)=1") & _
    "    Order by A.����"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not lvwMain.SelectedItem Is Nothing Then
        strKey = lvwMain.SelectedItem.Key
    End If
    lvwMain.ListItems.Clear
    Do While Not rsTemp.EOF
        If NVL(rsTemp!����) = "��" Then
            strͣ�� = "Start"
        Else
            strͣ�� = "Stop"
        End If
        Set objItem = lvwMain.ListItems.Add(, "K" & rsTemp!id, rsTemp!����, strͣ��, strͣ��)
        If strͣ�� = "Stop" Then objItem.ForeColor = RGB(255, 0, 0)
        objItem.Tag = IIf(NVL(rsTemp!����) = "��", 0, 1) & "-" & IIf(NVL(rsTemp!�̶���) = "��", 1, 0) & "-" & IIf(NVL(rsTemp!֤��) = "��", 1, 0)
        '����ListView�����������ݿ�ȡ��
        For lngCol = 2 To lvwMain.ColumnHeaders.count
            varValue = rsTemp(lvwMain.ColumnHeaders(lngCol).Text).value
            objItem.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
            If strͣ�� = "Stop" Then objItem.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
        Next
        rsTemp.MoveNext
    Loop
    
    If lvwMain.ListItems.count > 0 Then
        On Error Resume Next
        Set objItem = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Err.Clear
            Set objItem = lvwMain.ListItems(1)
            objItem.Selected = True
            objItem.EnsureVisible
        Else
            objItem.Selected = True
            objItem.EnsureVisible
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub InitLvwHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '����:���˺�
    '����:2011-06-28 00:48:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strText As String, i As Integer
    On Error GoTo Errhand
    lvwMain.Tag = "�ɱ仯��"
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
    If lvwMain.ColumnHeaders.count = 0 Then
        zlControl.LvwSelectColumns lvwMain, mstrLvw, True
    End If
    strText = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\zl9CardSquare\" & Me.Name & "\ListView", lvwMain.Name & "����")
    For i = 1 To lvwMain.ColumnHeaders.count
        '���������У��򲻻ָ����Ի�
        If InStr(strText, lvwMain.ColumnHeaders(i).Text) = 0 Then lvwMain.Tag = "": Exit For
        '����������У�Ҳ���ָ����Ի�
        strText = Replace(strText, lvwMain.ColumnHeaders(i).Text, "")
    Next
    strText = Replace(strText, ",", "")
    If strText <> "" Then lvwMain.Tag = ""
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub zlCardStopAndResume(Optional blnStop As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ͣ�û�����
    '����:���˺�
    '����:2011-06-27 20:56:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTypeId As Long, lngColor As Long, i As Long
    Dim strSQL As String, intIndex As Integer
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    'If Val(varTemp(1)) = 1 Then Exit Sub     'ϵͳ�̶���,������ͣ�ú�����
    Err = 0: On Error GoTo Errhand
    lngTypeId = Val(Mid(Me.lvwMain.SelectedItem.Key, 2))
    With lvwMain
         If MsgBox("��ȷ��Ҫ" & IIf(blnStop, "ͣ��", "����") & "ҽ�ƿ�""" & lvwMain.SelectedItem.Text & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            strSQL = "Zl_ҽ�ƿ����_Stopandstart(" & lngTypeId & "," & IIf(blnStop, 1, 0) & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            If mblnShowStop = False And blnStop Then
                    intIndex = .SelectedItem.Index
                    .ListItems.Remove .SelectedItem.Key
                    If .ListItems.count > 0 Then
                            intIndex = IIf(.ListItems.count > intIndex, intIndex, .ListItems.count)
                            .ListItems(intIndex).Selected = True
                            .ListItems(intIndex).EnsureVisible
                    Else
                        Call lvwMain_GotFocus
                    End If
            Else
                If blnStop Then
                    .SelectedItem.Icon = "Stop": .SelectedItem.SmallIcon = "Stop"
                    lngColor = vbRed
                Else
                    .SelectedItem.Icon = "Start": .SelectedItem.SmallIcon = "Start"
                    lngColor = RGB(0, 0, 0)
                End If
                .SelectedItem.ForeColor = lngColor
                For i = 1 To .ColumnHeaders.count
                    If i < .ColumnHeaders.count Then
                        .SelectedItem.ListSubItems(i).ForeColor = lngColor
                    End If
                    If .ColumnHeaders(i).Text = "�Ƿ�����" Then
                        .SelectedItem.SubItems(i - 1) = IIf(blnStop, "", "��")
                    End If
                Next
                .SelectedItem.Tag = IIf(blnStop, 1, 0) & "-" & varTemp(1) & "-" & varTemp(2)
            End If
        End If
    End With
    zlCtlSetFocus lvwMain, True
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
Private Sub ModifyData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸�����
    '����:���˺�
    '����:2011-06-28 00:59:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTypeId As Long, varTemp As Variant
    
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    If Val(varTemp(0)) = 1 Or Val(varTemp(2)) = 1 Then Exit Sub
'    If Val(varTemp(1)) = 1 Then
'        MsgBox "ϵͳ�̶���,�����޸�,����!", vbOKOnly + vbInformation, gstrSysName
'        Exit Sub
'    End If
    lngTypeId = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If frmPatiCureCardTypeEdit.zlEditCard(Me, mlngModule, mstrPrivs, edt_�޸�, lngTypeId) = False Then Exit Sub
    If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call LoadData
End Sub
Private Sub DeleteData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸�����
    '����:���˺�
    '����:2011-06-28 00:59:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTypeId As Long, varTemp As Variant
    Dim intIndex As Integer, strSQL As String
    Err = 0: On Error GoTo Errhand:
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    If Val(varTemp(0)) = 1 Or Val(varTemp(2)) = 1 Then Exit Sub
    If Val(varTemp(1)) = 1 Then
        MsgBox "ϵͳ�̶���,����ɾ��,����!", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvwMain.SelectedItem.Text & "����ҽ�ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    lngTypeId = Val(Mid(lvwMain.SelectedItem.Key, 2))
        'Zl_ҽ�ƿ����_Delete(Id_In In ҽ�ƿ����.ID%Type) Is
      strSQL = "Zl_ҽ�ƿ����_Delete(" & lngTypeId & ")"
      
      Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    With lvwMain
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.count > 0 Then
            intIndex = IIf(.ListItems.count > intIndex, intIndex, .ListItems.count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ViewData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ѯ����
    '����:���˺�
    '����:2011-06-28 00:59:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTypeId As Long
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    lngTypeId = Val(Mid(lvwMain.SelectedItem.Key, 2))
    Call frmPatiCureCardTypeEdit.zlEditCard(Me, mlngModule, mstrPrivs, dt_�鿴, lngTypeId)
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '���:
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-18 16:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objComBar As CommandBarComboBox
        
      
    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.id = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With


    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.id = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)"):
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ��(&S)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.id = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "��ͼ��(&L)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "Сͼ��(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "�б�(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "��ϸ����(&D)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "��ʾͣ����Ŀ(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.id = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long
    Dim ctrCombox As CommandBarComboBox
    '------------------------------------
        

    Select Case Control.id
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_NewItem
        If frmPatiCureCardTypeEdit.zlEditCard(Me, mlngModule, mstrPrivs, edT_����) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadData
    Case conMenu_Edit_Modify
        '�޸�
        Call ModifyData
    Case conMenu_Edit_Delete    'ɾ��
        Call DeleteData
    Case conMenu_Edit_Reuse  '����
          Call zlCardStopAndResume(False)
    Case conMenu_Edit_Stop 'ͣ��
          Call zlCardStopAndResume(True)
    Case conMenu_View_ShowStoped '��ʾͣ����ֹ
            mblnShowStop = Not mblnShowStop
            Call LoadData
    Case conMenu_View_LargeICO  '��ͼ��
         lvwMain.View = lvwIcon
    Case conMenu_View_MinICO    'Сͼ��
         lvwMain.View = lvwSmallIcon
    Case conMenu_View_ListICO   '�б�
         lvwMain.View = lvwList
    Case conMenu_View_DetailsICO    '��ϸ����
         lvwMain.View = lvwReport
    Case conMenu_View_Refresh   'ˢ��
        '����ˢ������
        Call LoadData
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
 
Private Function IsModifyOrDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ������޸Ļ�ɾ��
    '����:���˺�
    '����:2011-06-28 11:54:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    IsModifyOrDelete = Val(varTemp(0)) = 0 And Val(varTemp(1)) = 0 And Val(varTemp(2)) = 0
End Function
Private Function IsModify() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ������޸Ļ�ɾ��
    '����:���˺�
    '����:2011-06-28 11:54:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    ' �Ƿ�����-�Ƿ�̶�
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    IsModify = Val(varTemp(0)) = 0 And Val(varTemp(2)) = 0
End Function
Private Function IsStop() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ�����ͣ��
    '����:���˺�
    '����:2011-06-28 11:54:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    IsStop = Val(varTemp(0)) = 0 ' And Val(varTemp(1)) = 0
End Function
Private Function IsStart() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ���������
    '����:���˺�
    '����:2011-06-28 11:54:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    IsStart = Val(varTemp(0)) = 1 ' And Val(varTemp(1)) = 0
End Function

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngID As Long, blnEnabled As Boolean
     
    If Me.Visible = False Then Exit Sub
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = lvwMain.ListItems.count > 0
    Case conMenu_Edit_NewItem
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�޸�")
        Control.Enabled = Control.Visible And IsModify
    Case conMenu_Edit_Delete
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ɾ��")
        Control.Enabled = Control.Visible And IsModifyOrDelete
    Case conMenu_Edit_Reuse
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible And IsStart
    Case conMenu_Edit_Stop
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ͣ��")
        Control.Enabled = Control.Visible And IsStop
    Case conMenu_View_ShowStoped '��ʾͣ����ֹ
        Control.Checked = mblnShowStop
    Case conMenu_View_LargeICO  '��ͼ��
        Control.Checked = lvwMain.View = lvwIcon
    Case conMenu_View_MinICO    'Сͼ��
        Control.Checked = lvwMain.View = lvwSmallIcon
    Case conMenu_View_ListICO   '�б�
        Control.Checked = lvwMain.View = lvwList
    Case conMenu_View_DetailsICO    '��ϸ����
        Control.Checked = lvwMain.View = lvwReport
    Case conMenu_View_Refresh   'ˢ��
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_1" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1107_2"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.id
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else   '�����������ܵ���
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Resize()
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    cbsThis.GetClientRect Left, Top, Right, Bottom
    On Error Resume Next
   With lvwMain
        .Left = Left
        .Top = Top
        .Width = Right - Left
        .Height = Bottom - Top
   End With
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Load()
    mlngModule = glngModul
    mstrPrivs = gstrPrivs
    Call zlDefCommandBars
    Call InitLvwHead
     
    '76905,Ƚ����,2014-8-21,��һ�ν���ҽ�ƿ���������,��ͣ�������Ĭ����ʾʱ,δ��ʾͣ�����
    Call InitPara
    Call LoadData
    RestoreWinState Me, App.ProductName
    lvwMain.Tag = "�ɱ仯��"
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���������
    '����:���˺�
    '����:2011-06-28 11:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    mblnShowStop = Val(zlDatabase.GetPara("��ʾͣ�����", glngSys, mlngModule, "1", , , InStr(1, mstrPrivs, ";��������;") > 0)) = 1
    i = Val(zlDatabase.GetPara("ͼ����ʾ��ʽ", glngSys, mlngModule, "3", , , InStr(1, mstrPrivs, ";��������;") > 0))
    If i < 0 Or i > 3 Then i = 3
    lvwMain.View = i
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    Call zlDatabase.SetPara("��ʾͣ�����", IIf(mblnShowStop, 1, 0), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
    Call zlDatabase.SetPara("ͼ����ʾ��ʽ", lvwMain.View, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
 
End Sub

Private Sub lvwMain_DblClick()
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If lvwMain.SelectedItem.Tag Like "1*" Or lvwMain.SelectedItem.Tag Like "*-1" Or InStr(1, mstrPrivs, ";�޸�;") = 0 Then
        Call ViewData
    Else
        Call ModifyData
    End If
End Sub
Private Sub lvwMain_GotFocus()
    With lvwMain
        stbThis.Panels(2).Text = "����" & .ListItems.count & "��ҽ�ƿ����"
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:
    '����:
    '����:���˺�
    '����:2009-11-20 15:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = gstrUnitName & "ҽ������嵥"
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & UserInfo.����
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ������
    '���:lngSys-ϵͳ��
    '     strReportCode������
    '����:���˺�
    '����:2009-11-19 14:15:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID  As Long
    With lvwMain
        If Not .SelectedItem Is Nothing Then
            lngID = Val(Mid(.SelectedItem.Key, 2))
        End If
    End With
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, "ID=" & lngID)
End Sub

Private Sub zlPopuMenus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����˵�
    '����:���˺�
    '����:2011-06-28 12:18:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrPopupBar As CommandBar, cbrPopupItem As CommandBarControl
    
    Err = 0: On Error Resume Next
    If Not (Me.cbsThis.ActiveMenuBar.Controls(2).Visible Or Me.cbsThis.ActiveMenuBar.Controls(3).Visible) Then Exit Sub
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible Then
        Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
        For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
            If mcbrControl.id <> conMenu_View_ToolBar Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
            cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
            End If
        Next
    End If
    If Me.cbsThis.ActiveMenuBar.Controls(3).Visible Then
        Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(3)
        For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
            Select Case mcbrControl.id
            Case conMenu_View_ToolBar
            Case Else
                If mcbrControl.Caption Like "������*" Then
                Else
                    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
                    cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
                End If
            End Select
        Next
    End If
    cbrPopupBar.ShowPopup
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Exit Sub
    Call zlPopuMenus
End Sub
