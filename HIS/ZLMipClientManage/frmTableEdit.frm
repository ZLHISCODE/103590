VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmTableEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   8325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12165
   Icon            =   "frmTableEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   7860
      Index           =   0
      Left            =   30
      ScaleHeight     =   7860
      ScaleWidth      =   12195
      TabIndex        =   15
      Top             =   540
      Width           =   12195
      Begin VB.Frame fra 
         Height          =   7125
         Left            =   30
         TabIndex        =   16
         Top             =   -90
         Width           =   12105
         Begin VB.TextBox txt 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   1
            Left            =   1845
            TabIndex        =   1
            Text            =   "001"
            Top             =   285
            Width           =   720
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            ForeColor       =   &H80000010&
            Height          =   300
            Left            =   780
            TabIndex        =   18
            Text            =   "ZLHIS_USER_"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Frame Frame1 
            Height          =   4545
            Left            =   780
            TabIndex        =   17
            Top             =   555
            Width           =   11145
            Begin RichTextLib.RichTextBox txtSQL 
               Height          =   3930
               Left            =   90
               TabIndex        =   7
               Top             =   135
               Width           =   10905
               _ExtentX        =   19235
               _ExtentY        =   6932
               _Version        =   393217
               BorderStyle     =   0
               ScrollBars      =   1
               TextRTF         =   $"frmTableEdit.frx":000C
            End
            Begin VB.CommandButton cmdVerfiy 
               Caption         =   "У��(&V)"
               Height          =   350
               Left            =   1290
               TabIndex        =   9
               Top             =   4140
               Width           =   1100
            End
            Begin VB.CommandButton cmdPara 
               Caption         =   "����(&P)"
               Height          =   350
               Left            =   75
               TabIndex        =   8
               Top             =   4125
               Width           =   1100
            End
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   5580
            TabIndex        =   5
            Top             =   225
            Width           =   6345
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   3090
            TabIndex        =   3
            Top             =   240
            Width           =   1980
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1875
            Index           =   0
            Left            =   765
            TabIndex        =   11
            Top             =   5145
            Width           =   11160
            _cx             =   19685
            _cy             =   3307
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483638
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   270
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "SQL����"
            Height          =   180
            Index           =   5
            Left            =   90
            TabIndex        =   10
            Top             =   5160
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "SQL���"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   6
            Top             =   675
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "˵��"
            Height          =   180
            Index           =   2
            Left            =   5145
            TabIndex        =   4
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   2685
            TabIndex        =   2
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   0
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(C)"
         Height          =   350
         Left            =   10950
         TabIndex        =   13
         Top             =   7260
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9675
         TabIndex        =   12
         Top             =   7260
         Width           =   1100
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1470
      TabIndex        =   14
      Top             =   75
      Width           =   1575
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmTableEdit"
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
Private mintParsCount As Long
Private mblnContiune As Boolean
Private mlngModualCode As Long
Private mstrBusiness As String

Private WithEvents mclsVsf As zlVSFlexGrid.clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Event AfterNewData(ByVal DataKey As String)
Public Event AfterModifyData(ByVal DataKey As String)
Public Event AfterDeleteData(ByVal DataKey As String)
Public Event Forward(ByRef DataKey As String, ByRef Cancel As Boolean)
Public Event Backward(ByRef DataKey As String, ByRef Cancel As Boolean)

'######################################################################################################################

Public Function InitDialog(ByVal frmParent As Object, ByVal lngModualCode As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mlngModualCode = lngModualCode
    
    InitDialog = True
    
End Function

Public Sub NewData(ByVal strBusiness As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 1
    mstrDataKey = ""
    mstrBusiness = strBusiness
    
    Me.Caption = "����ҵ����Ϣ"
        
    Call InitData
    Call InitGrid
    Call InitCommandBar
    
    txt(1).Text = gclsBusiness.GetMaxUserTableCode("ZLHIS_USER_")
    
    mblnDataChanged = False
    
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub ModifyData(ByVal strBusiness As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mbytMode = 2
    mstrDataKey = strDataKey
    mstrBusiness = strBusiness
    
    Call InitData
    Call InitGrid
    Call InitCommandBar
    
    txt(1).Text = gclsBusiness.GetMaxCode("zlmip_table", "tab_code")
    
    Me.Caption = "�޸�ҵ����Ϣ"
        
    Call ReadData(mstrDataKey)
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub DeleteData(ByVal strBusiness As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 3
    mstrBusiness = strBusiness
    If strDataKey = "" Then Exit Sub
    mstrDataKey = strDataKey
    
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "ID", mstrDataKey)
        
    If gclsBusiness.TableEdit("Delete", mrsPara) Then
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
    
    Set rsTmp = gclsBusiness.TableStruct()
    If Not (rsTmp Is Nothing) Then
        txt(0).MaxLength = rsTmp("tab_title").DefinedSize
        txt(1).MaxLength = rsTmp("tab_code").Precision - Len(txtCode.Text)
        txt(2).MaxLength = rsTmp("tab_note").DefinedSize
        txtSQL.MaxLength = rsTmp("tab_sqltext").DefinedSize
    End If
    
    InitData = True
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '��ʼ����ؼ�
    Set mclsVsf = New zlVSFlexGrid.clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf(0), True, False, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 270, flexAlignLeftCenter, flexDTString, , "[���]", False, False, False)
        
        Call .AppendColumn("��������", 1500, flexAlignLeftCenter, flexDTString, , "para_title", True)
        Call .AppendColumn("��������", 1500, flexAlignLeftCenter, flexDTString, , "para_type", True)
        Call .AppendColumn("����ȱʡ", 0, flexAlignLeftCenter, flexDTString, , "para_default", True, , , True)
        Call .AppendColumn("����˵��", 3000, flexAlignLeftCenter, flexDTString, , "para_note", True)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        .UpdateSerial
        .AppendRows = True
        
        Call .InitializeEdit(True, False, False)
        
        Call .InitializeEditColumn(.ColIndex("��������"), True, vbVsfEditText)
        Call .InitializeEditColumn(.ColIndex("��������"), True, vbVsfEditCombox, "�ַ�|��ֵ|����")
        Call .InitializeEditColumn(.ColIndex("����ȱʡ"), True, vbVsfEditText)
        Call .InitializeEditColumn(.ColIndex("����˵��"), True, vbVsfEditText)
    End With
                
    InitGrid = True
    
End Function

Private Function ReadData(ByVal strDataKey As String) As Boolean

    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)
    
    mblnReading = True
    Set rsTmp = gclsBusiness.TableRead("id", rsCondition)
    If rsTmp.BOF = False Then
        txt(0).Text = zlCommFun.NVL(rsTmp("tab_title").Value)
        txt(1).Text = Replace(zlCommFun.NVL(rsTmp("tab_code").Value), "ZLHIS_USER_", "")
        txt(2).Text = zlCommFun.NVL(rsTmp("tab_note").Value)
        txtSQL.Text = zlCommFun.NVL(rsTmp("tab_sqltext").Value)
    End If
    
    Set rsTmp = gclsBusiness.TableParameterRead("tab_id", rsCondition)
    If rsTmp.BOF = False Then Call mclsVsf.LoadGrid(rsTmp)
    
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
    cbsMain.VisualTheme = xtpThemeWhidbey
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
        
    mstrFindKey = zlDataBase.GetPara("��λ����", ParamInfo.ϵͳ��, mlngModualCode, "����")
    If mstrFindKey = "" Then mstrFindKey = "����"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, True, , , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.����"): objControl.Parameter = "����"
    objControl.IconId = 1
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.����"): objControl.Parameter = "����"
    objControl.IconId = 1

    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "����")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Forward, "��һ��")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Backward, "��һ��")

    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, IIf(mbytMode = 1, "ȷ��֮��������", "ȷ��֮�����޸�"), False)
    objControl.IconId = conMenu_View_UnCheck
    If mbytMode <> 1 Then objControl.flags = xtpFlagRightAlign

    
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
    Dim lngLoop As Long
    
    If Len(txt(0).Text) = 0 Then
        ShowSimpleMsg "ҵ����Ϣ�����Ʋ���Ϊ�գ�"
        Call LocationObj(txt(0))
        Exit Function
    End If
    
    If Len(txt(1).Text) = 0 Then
        ShowSimpleMsg "ҵ����Ϣ�ı��벻��Ϊ�գ�"
        Call LocationObj(txt(1))
        Exit Function
    End If
    
    '�������Ƿ�Ϊ�����ַ�
    If zlCommFun.CheckStrType(Trim(txt(1).Text), 99, "0123456789") = False Then
        ShowSimpleMsg "�������Ϊ�����ַ���"
        LocationObj txt(1)
        Exit Function
    End If
    
    If Len(txtSQL.Text) = 0 Then
        ShowSimpleMsg "ҵ����Ϣ��SQL��䲻��Ϊ�գ�"
        Call LocationObj(txtSQL)
        Exit Function
    End If
        
    If VerfiySQL = False Then
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
    Dim intType As Integer
    Dim strLine As String
    Dim strTemp As String
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
    
    Call zlCommFun.SetParameter(rsPara, "data_code", mstrBusiness)
    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
    Call zlCommFun.SetParameter(rsPara, "tab_code", txtCode.Text & Trim(txt(1).Text))
    Call zlCommFun.SetParameter(rsPara, "tab_title", Trim(txt(0).Text))
    Call zlCommFun.SetParameter(rsPara, "tab_sqltext", Replace(txtSQL.Text, "'", "''"))
    Call zlCommFun.SetParameter(rsPara, "tab_note", Trim(txt(2).Text))
        
    '------------------------------------------------------------------------------------------------------------------
    With vsf(0)
        lngCount = 0
        strTemp = ""
        For lngLoop = 1 To .Rows - 1
            
            Select Case .TextMatrix(lngLoop, .ColIndex("��������"))
            Case "��ֵ"
                intType = 1
            Case "�ַ�"
                intType = 2
            Case "����"
                intType = 3
            End Select
            
            strLine = lngLoop
            
            strLine = strLine & "," & .TextMatrix(lngLoop, .ColIndex("��������"))
            strLine = strLine & "," & intType
            strLine = strLine & "," & .TextMatrix(lngLoop, .ColIndex("����ȱʡ"))
            strLine = strLine & "," & .TextMatrix(lngLoop, .ColIndex("����˵��"))
                        
            If LenB(strTemp & ";" & strLine) > 3500 Then
                If strTemp <> "" Then
                    lngCount = lngCount + 1
                    strTemp = Mid(strTemp, 2)
                    Call zlCommFun.SetParameter(rsPara, "SQL����_" & lngCount, strTemp)
                    strTemp = ""
                End If
            End If
            strTemp = strTemp & ";" & strLine
        Next
    End With
    If strTemp <> "" Then
        lngCount = lngCount + 1
        strTemp = Mid(strTemp, 2)
        Call zlCommFun.SetParameter(rsPara, "SQL����_" & lngCount, strTemp)
    End If
    Call zlCommFun.SetParameter(rsPara, "SQL��������", lngCount)
    
    
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
    strLine = ""
    lngCount = 0
    Set rsTmp = gclsBusiness.GetSQLField(Trim(txtSQL.Text))
    If Not (rsTmp Is Nothing) Then
        If rsTmp.BOF = False Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                strLine = rsTmp("���").Value
                strLine = strLine & "," & rsTmp("����").Value
                strLine = strLine & "," & rsTmp("����").Value
                
                If LenB(strTemp & ";" & strLine) > 3500 Then
                    If strTemp <> "" Then
                        lngCount = lngCount + 1
                        strTemp = Mid(strTemp, 2)
                        Call zlCommFun.SetParameter(rsPara, "SQL�ֶ�_" & lngCount, strTemp)
                        strTemp = ""
                    End If
                End If
                strTemp = strTemp & ";" & strLine
            
                rsTmp.MoveNext
            Loop
        End If
    Else
        GoTo errHand
    End If
    
    If strTemp <> "" Then
        lngCount = lngCount + 1
        strTemp = Mid(strTemp, 2)
        Call zlCommFun.SetParameter(rsPara, "SQL�ֶ�_" & lngCount, strTemp)
    End If
    Call zlCommFun.SetParameter(rsPara, "SQL�ֶθ���", lngCount)
    
    
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '����
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        
        SaveData = gclsBusiness.TableEdit("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�޸�
        SaveData = gclsBusiness.TableEdit("UPDATE", rsPara)
    End Select
    
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

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
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        
        Dim strText As String
        Dim rsCondition As ADODB.Recordset
        Dim rsData As ADODB.Recordset
        Dim rs As ADODB.Recordset
        
        If txtLocation.Text <> "" Then
            
            txtLocation.Tag = ""
                        
            Set rsCondition = zlCommFun.CreateCondition
            
            Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
            Call zlCommFun.SetCondition(rsCondition, "FilterText", txtLocation.Text)
            Set rsData = gclsBusiness.TableRead("FilterData", rsCondition)
                        
            If zlCommFun.ShowPubSelect(Me, txtLocation, 2, "����,900,0,1;����,1800,0,0;˵��,2400,0,0", Me.Name & "\ҵ����Ϣ�����", "����±���ѡ��һ��ҵ����Ϣ��", rsData, rs, , , , , , True) = 1 Then
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
                        txt(1).Text = gclsBusiness.GetMaxUserTableCode("ZLHIS_USER_")
                        txt(2).Text = ""
                        txtSQL.Text = ""
                    End If
                    Call LocationObj(txt(0))
                    mblnDataChanged = False
                End If
            End If
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub cmdPara_Click()
    Dim strSQL As String
    Dim intLoop As Integer
    Dim rsSQLPara As ADODB.Recordset
    
    strSQL = TrimChar(RemoveNote(txtSQL.Text))
        
    If strSQL <> "" Then
            
        '����������
        '------------------------------------------------------------------------------------------------------------------
        If gclsBusiness.CheckSQLPara(strSQL) = False Then
            MsgBox "�������岻��ȷ��������������Ƿ����,�������Ƿ�Ϊ��ֵ,��������ţ�", vbInformation + vbOKOnly, ParamInfo.ϵͳ����
            txtSQL.SetFocus
            Exit Sub
        End If
                
        Set rsSQLPara = gclsBusiness.GetSQLPara(strSQL)
                
        If rsSQLPara Is Nothing Then
            mclsVsf.ClearGrid
        Else
            If rsSQLPara.RecordCount = 0 Then
                mclsVsf.ClearGrid
            Else
                
                With vsf(0)
                    .Rows = rsSQLPara.RecordCount + 1
                    For intLoop = 1 To .Rows - 1
                        
                        If .TextMatrix(intLoop, .ColIndex("��������")) = "" Then
                            .TextMatrix(intLoop, .ColIndex("��������")) = "����" & intLoop
                            .TextMatrix(intLoop, .ColIndex("��������")) = "�ַ�"
                        End If
                        
                    Next
                    mclsVsf.UpdateSerial
                    mclsVsf.AppendRows = True
                End With
        
            End If
        End If
    End If
        
End Sub

Private Sub cmdVerfiy_Click()
    
    If VerfiySQL Then
        MsgBox "��ǰSQL�ǺϷ��ģ�", vbInformation, Me.Caption
    End If
    
End Sub

Private Function VerfiySQL() As Boolean
    Dim intLoop As Integer
    Dim rsSQLPara As ADODB.Recordset
            
    '------------------------------------------------------------------------------------------------------------------
    Set rsSQLPara = New ADODB.Recordset
    With rsSQLPara
        .Fields.Append "���", adTinyInt
        .Fields.Append "����", adVarChar, 60
        .Fields.Append "����", adVarChar, 10
        .Open
    End With
    With vsf(0)
        For intLoop = 1 To .Rows - 1
            If .TextMatrix(intLoop, .ColIndex("��������")) <> "" Then
                rsSQLPara.AddNew
                rsSQLPara("���").Value = intLoop - 1
                rsSQLPara("����").Value = .TextMatrix(intLoop, .ColIndex("��������"))
                rsSQLPara("����").Value = .TextMatrix(intLoop, .ColIndex("��������"))
            End If
        Next
    End With
    
    VerfiySQL = gclsBusiness.CheckSQL(TrimChar(RemoveNote(txtSQL.Text)), rsSQLPara)
    
End Function

Private Function RemoveNote(ByVal strSQL As String) As String
    '���ܣ��Ƴ�SQL����е�ע��
    '˵����ֻ֧���Ƴ����е�ע��
    Dim strTmp As String
    Dim i As Integer
    Dim arrLine() As String
    
    strSQL = Replace(strSQL, vbTab, " ")
    strSQL = Replace(strSQL, vbLf, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr, vbCrLf)
    
'    strSQL = Replace(strSQL, "'", "''")
    
    arrLine = Split(strSQL, vbCrLf)
    
    For i = 0 To UBound(arrLine)
        If Not Trim(arrLine(i)) Like "--*" Then
            RemoveNote = RemoveNote & vbCrLf & arrLine(i)
        End If
    Next
    RemoveNote = Mid(RemoveNote, 3)
End Function

Private Function TrimChar(ByVal strSQL As String) As String
'����:ȥ���ַ����������Ŀո�ͻس�(����ͷ�Ŀո�,�س�),��ȥ��TAB�ַ�,������������
    Dim strTmp As String
    Dim i As Long
    Dim j As Long
    
    If Trim(strSQL) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(strSQL)
    
    strTmp = Replace(strTmp, "  ", " ")
    strTmp = Replace(strTmp, "  ", " ")
    
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)

    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Set mclsVsf = Nothing
    
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = vsf(0).Rows > mintParsCount
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    mblnDataChanged = True
        
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    Select Case Index
    Case 4
        
    Case Else
        zlControl.TxtSelAll txt(Index)
    End Select
    
    Select Case Index
    Case 0, 2, 4
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

        Select Case Index
        Case 1
            If zlCommFun.FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        End Select
        
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

Private Sub txtSQL_Change()
    If mblnReading Then Exit Sub
    mblnDataChanged = True
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
    If mblnReading Then Exit Sub
    mblnDataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
    End Select
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub
