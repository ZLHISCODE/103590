VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQueryFilter 
   Caption         =   "��ѯ����"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5625
   Icon            =   "frmQueryFilter.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5460
   ScaleWidth      =   5625
   StartUpPosition =   1  '����������
   Begin MSComctlLib.Slider sdrRange 
      Height          =   230
      Left            =   2040
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   3132
      _ExtentX        =   5530
      _ExtentY        =   397
      _Version        =   393216
      Max             =   366
      TickStyle       =   3
      TickFrequency   =   3
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "���(&C)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4920
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ ��(&Q)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4296
      TabIndex        =   14
      Top             =   4920
      Width           =   1185
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2928
      TabIndex        =   13
      Top             =   4920
      Width           =   1185
   End
   Begin VB.Frame framButton 
      Height          =   795
      Left            =   -120
      TabIndex        =   9
      Top             =   4680
      Width           =   5895
   End
   Begin VB.ListBox lstObj 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Index           =   0
      ItemData        =   "frmQueryFilter.frx":000C
      Left            =   2088
      List            =   "frmQueryFilter.frx":000E
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   2496
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cbxObj 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   0
      Left            =   2088
      TabIndex        =   7
      Text            =   "cbxObj"
      Top             =   2100
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtObj 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2088
      TabIndex        =   6
      Top             =   1668
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cbxWhere 
      BackColor       =   &H8000000F&
      Height          =   276
      ItemData        =   "frmQueryFilter.frx":0010
      Left            =   645
      List            =   "frmQueryFilter.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1620
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkObj 
      Caption         =   "��ѡ����"
      Height          =   255
      Index           =   0
      Left            =   2088
      TabIndex        =   3
      Top             =   4020
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ComboBox cbxAge 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   0
      ItemData        =   "frmQueryFilter.frx":0014
      Left            =   4365
      List            =   "frmQueryFilter.frx":0024
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4380
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAge 
      Height          =   345
      Index           =   0
      Left            =   2088
      MaxLength       =   3
      TabIndex        =   1
      Top             =   4380
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cbxDateUnit 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   0
      ItemData        =   "frmQueryFilter.frx":0040
      Left            =   4125
      List            =   "frmQueryFilter.frx":005F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   1015
   End
   Begin MSScriptControlCtl.ScriptControl sctExecute 
      Left            =   900
      Top             =   3390
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComCtl2.DTPicker dtpObj 
      Height          =   324
      Index           =   0
      Left            =   2088
      TabIndex        =   5
      Top             =   1020
      Visible         =   0   'False
      Width           =   2052
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   55902209
      CurrentDate     =   41297
   End
   Begin VB.Label labObj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ռλ:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label labError 
      Alignment       =   2  'Center
      Caption         =   "û����Ҫ¼�����Ŀ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1470
      TabIndex        =   11
      Top             =   2505
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label labMemo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Left            =   792
      TabIndex        =   10
      Top             =   108
      Width           =   4656
   End
   Begin VB.Image imgQuery 
      Height          =   720
      Left            =   36
      Picture         =   "frmQueryFilter.frx":00A9
      Stretch         =   -1  'True
      Top             =   72
      Width           =   720
   End
   Begin VB.Shape shpBack 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CEFFFA&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5670
   End
End
Attribute VB_Name = "frmQueryFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Sql��ѯ����еĿ�ѡ��������Ϊ"[@������,�ֶ�='value']"

Private mobjFilterValue As clsSqlFilterValue
Private mobjSchemeItem As TSchemeItem

Private mobjLastControl As Object       '����¼�����ʱ���������һ�δ�����¼�����
Private maryInputTag() As TInputTag  '����¼�����ֵ�ı������������Ŀؼ�

Private mdblFontSize As Double
Private mdblZoomRate As Double
Private mblnIsLoading As Boolean
Private mblnIsOK As Boolean
Private mblnIsMoreInput As Boolean  '�Ƿ��и���¼����

Private mblnIsEmbed As Boolean      '�Ƿ�Ƕ��ʽ

Private maryItemData(50, 100) As String '����itemdata������

Private WithEvents mobjSqlParse As clsSqlParse
Attribute mobjSqlParse.VB_VarHelpID = -1


Property Get IsMoreInput() As Boolean
    IsMoreInput = mblnIsMoreInput
End Property

Property Get IsEmbed() As Boolean
    IsEmbed = mblnIsEmbed
End Property

Property Let IsEmbed(Value As Boolean)
    mblnIsEmbed = Value
End Property




Public Function ShowFilter(ByRef objSchemeItem As TSchemeItem, _
    ByVal dblFontSize As Double, owner As Object) As Boolean
'��ʾ���˴���
    ShowFilter = False
    
    If objSchemeItem.FilterValues Is Nothing Then
        Set objSchemeItem.FilterValues = New clsSqlFilterValue
    End If
    
    Set mobjFilterValue = objSchemeItem.FilterValues
    
    mobjSchemeItem = objSchemeItem
    
    If mobjSchemeItem.SqlScheme Is Nothing Then
        MsgBox "��ѯ���� [" & objSchemeItem.BaseInfo.Name & "] ��Ч�����ܻ�ȡ�����Ľ�������", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    Call ApplyOwnerFontSize(owner)
    
    If dblFontSize > 0 Then mdblFontSize = dblFontSize
    
    If mblnIsEmbed Then
        mblnIsOK = True
    Else
        mblnIsOK = False
        Me.Show 1, owner
    End If
    
    objSchemeItem = mobjSchemeItem
    
    ShowFilter = mblnIsOK
    
End Function

Public Sub UpdateInputData(ByVal strFilterName As String, strValue As Variant)
'���ý���¼������
    Dim inputTag As TInputTag
    Dim objInputControl As Object
    Dim i As Long
    Dim j As Long
    Dim strParName As String
    
    For i = 1 To UBound(maryInputTag)
        inputTag = maryInputTag(i)
        
        strParName = IIf(inputTag.ControlType = ctChk, "@", "") & inputTag.ParName
        
        If strParName = strFilterName Then
            Set objInputControl = inputTag.InputControl
            
            Select Case inputTag.ControlType
                Case ctText '�ı���
                    If inputTag.LikeWay <> lwNormal Then
                        objInputControl.Text = Replace(strValue, "%", "")
                    Else
                        objInputControl.Text = strValue
                    End If
                    
                Case ctDate, ctDateTime, ctTime, ctFastDate '���ڿ�
                    objInputControl.Value = strValue
                    
                    If sdrRange.Tag <> "" Then
                        If inputTag.Index = sdrRange.Tag Or inputTag.Index = sdrRange.Tag + 1 Then
                            sdrRange.Value = dtpObj(sdrRange.Tag + 1).Value - dtpObj(sdrRange.Tag).Value
                        End If
                    End If
                Case ctAgeCbx   '�����
                    If strValue = "" Then Exit Sub
                    
                    Select Case cbxAge(inputTag.Index)
                        Case "S-��"
                            strValue = CInt(Val(strValue) / 365) + IIf(Val(strValue) Mod 365 > 0, 1, 0)
                        Case "Y-��"
                            strValue = CInt(Val(strValue) / 30) + IIf(Val(strValue) Mod 30 > 0, 1, 0)
                        Case "Z-��"
                            strValue = CInt(Val(strValue) / 7) + IIf(Val(strValue) Mod 7 > 0, 1, 0)
                    End Select
                    
                    objInputControl.Text = strValue
                    
                Case ctCombobox '������
                    '�������������Ϊ0,��ʾû�н���ѡ��
                    If maryItemData(inputTag.Index, 1) <> "" Then
                        For j = 0 To objInputControl.ListCount - 1
                            If maryItemData(inputTag.Index, j) = strValue Then
                                inputTag.InputControl.ListIndex = j
                                Exit Sub
                            End If
                        Next j
                    Else
                        objInputControl.Text = strValue
                    End If
                    
                Case ctList '�б��
                    If maryItemData(inputTag.Index, 0) <> "" Then
                        For j = 0 To objInputControl.ListCount - 1
                            If InStr("," & strValue & ",", "," & maryItemData(inputTag.Index, j) & ",") > 0 Then
                                objInputControl.Selected(j) = True
                            End If
                        Next j
                    Else
                        For j = 0 To objInputControl.ListCount - 1
                            If InStr("," & strValue & ",", "," & objInputControl.List(j) & ",") > 0 Then
                                objInputControl.Selected(j) = True
                            End If
                        Next j
                    End If
                    
                Case ctChk  '��ѡ��
                    If CBool(strValue) <> False Then objInputControl.Value = 1
                    
                Case ctMutxCbx
                    If Trim(strValue) <> "" Then
                        cbxWhere.Text = inputTag.ParName
                        txtObj(cbxWhere.Tag).Text = strValue
                    End If
            End Select
            
            '��Ҫ����change�¼�
            Call ControlChange(inputTag, IIf(Trim(strValue) = "", True, False))
                    
            Exit Sub
        End If
    Next i
End Sub

Private Sub ApplyOwnerFontSize(owner As Object)
On Error GoTo errHandle
    Dim dblSize As Double
    
    dblSize = owner.FontSize
    
    mdblFontSize = dblSize
    
Exit Sub
errHandle:
End Sub

Public Sub SetFontSize(ByVal lngFontSize As Double)
'���������С
On Error Resume Next
    Dim i As Long
    Dim objControl As control
    
    If lngFontSize <= 0 Then Exit Sub
    
    For Each objControl In Me.Controls
        If objControl.Name <> "labMemo" Then objControl.Font.Size = lngFontSize
    Next
    
Err.Clear
End Sub

Private Sub ConfigTitleDisplay()
    Me.Caption = mobjSchemeItem.SqlScheme.SchemeName
    labMemo.Caption = "˵��:" & mobjSchemeItem.SqlScheme.Descript
End Sub

Private Function IsSql(ByVal strFrom As String) As Boolean
'�Ƿ�sql���
    Dim lngSelectIndex As Long
    Dim lngFromIndex As Long
    Dim strUCase As String
    
    IsSql = False
    strUCase = UCase(strFrom)
    
    lngSelectIndex = InStr(strUCase, "SELECT")
    lngFromIndex = InStr(strUCase, "FROM")
    
    If lngSelectIndex < 0 Or lngFromIndex < 0 Then Exit Function
    
    If lngFromIndex <= lngSelectIndex Then Exit Function
    
    IsSql = True
End Function


Private Sub ConfigSysDateInput(ByRef lngStartInputIndex As Long)
'����ϵͳʱ��¼��
    Dim inputTag As TInputTag
    
    '��ʼ������������
    inputTag.ParName = "ϵͳ.��ʼ����"
    inputTag.DisplayName = "[��ʼ����]"
    inputTag.DataFrom = ""
    inputTag.FromType = dbftText
    inputTag.ControlType = 1
    inputTag.Index = lngStartInputIndex
    inputTag.Default = mobjSchemeItem.Startdate
    
    ReDim inputTag.ParList(0)
    ReDim inputTag.ReleationInputIndex(0)
    
    Set inputTag.InputControl = CreateInputControl(inputTag.DisplayName, inputTag.ControlType, lngStartInputIndex, mobjSchemeItem.Startdate)

    
    ReDim Preserve maryInputTag(lngStartInputIndex)
    maryInputTag(lngStartInputIndex) = inputTag
    
    lngStartInputIndex = lngStartInputIndex + 1
    
    '����������������
    inputTag.ParName = "ϵͳ.��������"
    inputTag.DisplayName = "[��������]"
    inputTag.DataFrom = ""
    inputTag.FromType = dbftText
    inputTag.ControlType = 1
    inputTag.Index = lngStartInputIndex
    inputTag.Default = mobjSchemeItem.EndDate
    
    ReDim inputTag.ParList(0)
    ReDim inputTag.ReleationInputIndex(0)
    
    Set inputTag.InputControl = CreateInputControl(inputTag.DisplayName, inputTag.ControlType, lngStartInputIndex, mobjSchemeItem.EndDate)

    
    ReDim Preserve maryInputTag(lngStartInputIndex)
    maryInputTag(lngStartInputIndex) = inputTag
    
    lngStartInputIndex = lngStartInputIndex + 1
    
    sdrRange.Top = inputTag.InputControl.Top + inputTag.InputControl.Height + 20
    sdrRange.Left = inputTag.InputControl.Left
    sdrRange.Width = inputTag.InputControl.Width
    sdrRange.Tag = 1
    
    sdrRange.Value = CDate(Format(mobjSchemeItem.Startdate, "yyyy-MM-dd")) - CDate(Format(mobjSchemeItem.EndDate, "yyyy-MM-dd"))
    sdrRange.Visible = True
    
    Set mobjLastControl = sdrRange
End Sub


Private Sub ConfigInputControl()
'���ý���¼��
    Dim i As Long
    Dim objSqlScheme As clsSqlScheme
    Dim strParName As String
    Dim lngLastOrder As Long
    Dim inputTag As TInputTag
    Dim objSerachCfg As clsScSerachCfg
    Dim objSqlParse As clsSqlParse
    Dim lngInputIndex As Long
    
    lngInputIndex = 1
    
    ReDim maryInputTag(0)
    
    Set objSqlScheme = mobjSchemeItem.SqlScheme
    Set objSqlParse = New clsSqlParse
    
    '�ж��Ƿ���ϵͳ�Ŀ�ʼ���ںͽ�����������......
    If InStr(objSqlScheme.Query, "[ϵͳ.��ʼ����]") > 0 _
        And InStr(objSqlScheme.Query, "[ϵͳ.��������]") > 0 Then
        
        Call ConfigSysDateInput(lngInputIndex)
    End If
    '
    
    For i = 1 To objSqlScheme.SerachCfgCount
        Set objSerachCfg = objSqlScheme.SerachCfg(i)
        
        inputTag.ParName = objSerachCfg.Name
        inputTag.LikeWay = objSerachCfg.LikeWay
        inputTag.DataFrom = Trim$(objSerachCfg.DataFrom)
        inputTag.FromType = dbftText
        inputTag.ControlType = objSerachCfg.ControlType
        inputTag.Index = lngInputIndex
        inputTag.Default = objSerachCfg.Default
        
        '�ж�������Դ����
        If inputTag.DataFrom <> "" Then
            If IsSql(inputTag.DataFrom) Then
                inputTag.FromType = dbftSql
            End If
        End If
        
        ReDim inputTag.ParList(0)
        ReDim inputTag.ReleationInputIndex(0)
        
        If inputTag.FromType = 1 Then
            objSqlParse.init inputTag.DataFrom
            If objSqlParse.SqlStruct.ParCount > 0 Then
                CopyStrArray objSqlParse.SqlStruct.AllParameterAry, inputTag.ParList
            End If
        End If
        
        Set inputTag.InputControl = CreateInputControl(inputTag.ParName, inputTag.ControlType, lngInputIndex)
   
        ReDim Preserve maryInputTag(lngInputIndex)
        maryInputTag(lngInputIndex) = inputTag
        
        lngInputIndex = lngInputIndex + 1
    Next i
    
End Sub

Private Sub ControlChange(ByRef inputTag As TInputTag, Optional ByVal blnIsNull As Boolean = False)
'��ǰ�ؼ����ݸı��ͬ�������ؼ���������Դ�Ե�ǰ�ؼ���Ϊ����������
    Dim i As Long
    Dim j As Long
    Dim releationInputTag As TInputTag
    Dim lngBound As Long
    
    '���tagΪ�գ���������Ŀ������¼������
    If inputTag.Tag = "" Then
        ReDim inputTag.ReleationInputIndex(0)
        For i = inputTag.Index + 1 To UBound(maryInputTag)
            releationInputTag = maryInputTag(i)
            For j = 1 To UBound(releationInputTag.ParList)
                If releationInputTag.ParList(j) = "[" & inputTag.ParName & "]" Then
                
                    lngBound = UBound(inputTag.ReleationInputIndex) + 1
                    ReDim Preserve inputTag.ReleationInputIndex(lngBound)
                    
                    inputTag.ReleationInputIndex(lngBound) = i
                    Exit For
                End If
            Next j
        Next i
        
        inputTag.Tag = "1"
    End If
    
'    If blnIsNull Then
'        For i = 1 To UBound(inputTag.ReleationInputIndex)
'            Call ClearControlValue(maryInputTag(inputTag.ReleationInputIndex(i)).InputControl, maryInputTag(inputTag.ReleationInputIndex(i)).ControlType)
'        Next i
'    Else
        For i = 1 To UBound(inputTag.ReleationInputIndex)
            Call ConfigControlValue(maryInputTag(inputTag.ReleationInputIndex(i)), False)
        Next i
'    End If
    
End Sub


Private Function CreateInputControl(ByVal strName As String, ByVal lngInputType As Long, _
    ByVal lngOrder As Long, Optional ByVal strDefault As String = "") As Object
'����¼�����
    Dim lngChkObjCount As Long
    Dim lngStartLeft As Long
    Dim blnIsOption As Boolean
    Dim lngStartTop As Long
    
    lngStartLeft = 1750 '1950
    lngStartLeft = lngStartLeft * mdblZoomRate
    
    lngStartTop = IIf(mblnIsEmbed, 120, 1080)
    
    blnIsOption = False
    
    Set CreateInputControl = Nothing
    
    Select Case lngInputType
        Case ctText
            '�����ı������
            Load txtObj(lngOrder)
            
            txtObj(lngOrder).Tag = strName
            
            txtObj(lngOrder).Left = lngStartLeft
            
            If mobjLastControl Is Nothing Then
                txtObj(lngOrder).Top = lngStartTop '315
            Else
                txtObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            Set mobjLastControl = txtObj(lngOrder)
            
        Case ctDate, ctTime, ctDateTime, ctFastDate
            '�������ڿ����
            Load dtpObj(lngOrder)
                        
            dtpObj(lngOrder).Height = 288 * mdblZoomRate
            dtpObj(lngOrder).Format = dtpCustom
            dtpObj(lngOrder).CustomFormat = IIf(lngInputType = ctDate Or lngInputType = ctFastDate, "yyyy-MM-dd", IIf(lngInputType = ctTime, "HH:mm", "yyyy-MM-dd HH:mm"))
            
            dtpObj(lngOrder).UpDown = IIf(lngInputType = ctTime, True, False)
            
            dtpObj(lngOrder).Value = CurServerDate
            If strDefault <> "" Then dtpObj(lngOrder).Value = CDate(strDefault)

            dtpObj(lngOrder).Tag = strName
            
            dtpObj(lngOrder).Left = lngStartLeft
            
            
            If mobjLastControl Is Nothing Then
                dtpObj(lngOrder).Top = lngStartTop '315
            Else
                dtpObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
                        
            If lngInputType = ctFastDate Then
                '���ӿ�ѡ�ؼ�
                Load cbxDateUnit(lngOrder)
                
                cbxDateUnit(lngOrder).Tag = strName
                
                Call cbxDateUnit(lngOrder).AddItem("����")
                Call cbxDateUnit(lngOrder).AddItem("ǰһ��")
                Call cbxDateUnit(lngOrder).AddItem("ǰ����")
                Call cbxDateUnit(lngOrder).AddItem("ǰ����")
                Call cbxDateUnit(lngOrder).AddItem("ǰһ��")
                Call cbxDateUnit(lngOrder).AddItem("ǰ����")
                Call cbxDateUnit(lngOrder).AddItem("ǰһ��")
                Call cbxDateUnit(lngOrder).AddItem("ǰ����")
                Call cbxDateUnit(lngOrder).AddItem("ǰ����")
            
                cbxDateUnit(lngOrder).ListIndex = 0
                
                cbxDateUnit(lngOrder).Left = dtpObj(lngOrder).Left + (dtpObj(lngOrder).Width * mdblZoomRate) + 60
                cbxDateUnit(lngOrder).Width = cbxDateUnit(lngOrder).Width * mdblZoomRate
                cbxDateUnit(lngOrder).Top = dtpObj(lngOrder).Top
                
                cbxDateUnit(lngOrder).Visible = True
            Else
                dtpObj(lngOrder).Width = 3135
                
                If lngInputType = ctFastDate Then dtpObj(lngOrder).CheckBox = True
            End If
            
            Set mobjLastControl = dtpObj(lngOrder)
            
        Case ctCombobox
            '����������
            Load cbxObj(lngOrder)
            
            cbxObj(lngOrder).Tag = strName
            
            cbxObj(lngOrder).Left = lngStartLeft
            
            cbxObj(lngOrder).Text = ""
            
            If mobjLastControl Is Nothing Then
                cbxObj(lngOrder).Top = lngStartTop '315
            Else
                cbxObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            Set mobjLastControl = cbxObj(lngOrder)
        Case ctList
            '�����ɶ�ѡ���б��
            Load lstObj(lngOrder)
            
            lstObj(lngOrder).Height = 1400 * mdblZoomRate
            
            lstObj(lngOrder).Tag = strName
            
            lstObj(lngOrder).Left = lngStartLeft
            
            If mobjLastControl Is Nothing Then
                lstObj(lngOrder).Top = lngStartTop '315
            Else
                lstObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            Set mobjLastControl = lstObj(lngOrder)


        Case ctAgeCbx
            '������������
            Load txtAge(lngOrder)
            Load cbxAge(lngOrder)
            
            txtAge(lngOrder).Tag = strName
            cbxAge(lngOrder).Tag = strName
            
            txtAge(lngOrder).Left = lngStartLeft
            cbxAge(lngOrder).Left = lngStartLeft + (txtAge(lngOrder).Width * mdblZoomRate)
            cbxAge(lngOrder).Width = cbxAge(lngOrder).Width * mdblZoomRate
            
            If mobjLastControl Is Nothing Then
                txtAge(lngOrder).Top = lngStartTop '315
                cbxAge(lngOrder).Top = lngStartTop
            Else
                txtAge(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
                cbxAge(lngOrder).Top = txtAge(lngOrder).Top
            End If
            
            Call cbxAge(lngOrder).AddItem("S-��")
            Call cbxAge(lngOrder).AddItem("Y-��")
            Call cbxAge(lngOrder).AddItem("Z-��")
            Call cbxAge(lngOrder).AddItem("T-��")
            
            cbxAge(lngOrder).ListIndex = 0
            cbxAge(lngOrder).Visible = True
            
            Set mobjLastControl = txtAge(lngOrder)
            
        Case ctMutxCbx  '����������
            If Trim(cbxWhere.Tag) = "" Then
                Load txtObj(lngOrder)
                
                txtObj(lngOrder).Width = txtObj(lngOrder).Width * mdblZoomRate
'                txtObj(lngOrder).Tag = strName
                txtObj(lngOrder).Left = lngStartLeft
                
                If mobjLastControl Is Nothing Then
                    txtObj(lngOrder).Top = lngStartTop '315
                Else
                    txtObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
                End If
                                
                
                Set mobjLastControl = txtObj(lngOrder)
            Else
                
            End If
            
        Case ctChk '��ѡ����
            Load chkObj(lngOrder)
            chkObj(lngOrder).Tag = strName
            chkObj(lngOrder).Caption = strName
            
            chkObj(lngOrder).Width = TextWidth(strName) * 1.2 + 252
            
            If Val(strDefault) <> 0 Then
                chkObj(lngOrder).Value = 1
            End If
            
            chkObj(lngOrder).Left = lngStartLeft
            
            If mobjLastControl Is Nothing Then
                chkObj(lngOrder).Top = lngStartTop '315
            Else
                chkObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
'            lngChkObjCount = chkObj.Count
'            If (lngChkObjCount Mod 2) = 0 Then
'                chkObj(lngOrder).Left = lngStartLeft
'
'                If mobjLastControl Is Nothing Then
'                    chkObj(lngOrder).Top = 1080 '315
'                Else
'                    chkObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
'                End If
'            Else
'                If chkObj(chkObj.UBound - 1).Width > 1485 Then
'                    chkObj(lngOrder).Left = lngStartLeft
'                    chkObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
'                Else
'                    chkObj(lngOrder).Left = 3600
'                    chkObj(lngOrder).Top = chkObj(chkObj.UBound - 1).Top
'                End If
'            End If
            
            
            Set mobjLastControl = chkObj(lngOrder)
'            mobjLastControl.Visible = True
            
            blnIsOption = True
'            Exit Function

            
    End Select
    
    mobjLastControl.Visible = True
    Set CreateInputControl = mobjLastControl
    
    If blnIsOption Then
        Exit Function
    End If
    
    If lngInputType = ctMutxCbx Then
        If Trim(cbxWhere.Tag) = "" Then
            cbxWhere.Visible = True
            cbxWhere.Width = cbxWhere.Width * mdblZoomRate
            cbxWhere.Tag = lngOrder
            
            
            cbxWhere.Left = mobjLastControl.Left - cbxWhere.Width - 120
            cbxWhere.Top = mobjLastControl.Top + 30
        End If
        
        cbxWhere.AddItem strName
        cbxWhere.ListIndex = 0
        
        Set CreateInputControl = cbxWhere 'txtObj(cbxWhere.Tag)
    Else
        mobjLastControl.Width = mobjLastControl.Width * mdblZoomRate
        
        '����Label����
        Load labObj(lngOrder)
        
        labObj(lngOrder).Caption = strName
        labObj(lngOrder).Left = mobjLastControl.Left - labObj(lngOrder).Width - 120
        labObj(lngOrder).Top = mobjLastControl.Top + 60
        labObj(lngOrder).Visible = True
    End If
End Function

    
Private Sub cbxDateUnit_Change(Index As Integer)
On Error GoTo errHandle
    'dtpObj(Index).Value =
    Select Case cbxDateUnit(Index).ListIndex
        Case 0  '����
            dtpObj(Index).Value = CurServerDate
        Case 1  'ǰһ��
            dtpObj(Index).Value = CurServerDate - 1
        Case 2  'ǰ����
            dtpObj(Index).Value = CurServerDate - 2
        Case 3  'ǰ����
            dtpObj(Index).Value = CurServerDate - 3
        Case 4  'ǰһ��
            dtpObj(Index).Value = CurServerDate - 7
        Case 5  'ǰ����
            dtpObj(Index).Value = CurServerDate - 14
        Case 6  'ǰһ��
            dtpObj(Index).Value = CurServerDate - 30
        Case 7  'ǰ����
            dtpObj(Index).Value = CurServerDate - 90
        Case 8  'ǰ����
            dtpObj(Index).Value = CurServerDate - 180
    End Select
        
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxObj_Change(Index As Integer)
'����������ֵ���û��ı����Ҫ��������ݼ���
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index), IIf(cbxObj(Index).Text = "", True, False))
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxObj_Click(Index As Integer)
'����������ֵ���û��ı����Ҫ��������ݼ���
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index), IIf(cbxObj(Index).Text = "", True, False))
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub



Private Sub cmdCancel_Click()
On Error GoTo errHandle
    
    Me.Hide
    Exit Sub
errHandle:
End Sub

Private Sub cmdClear_Click()
mblnIsLoading = True

On Error GoTo errHandle
    Dim objFree As Object
    
    '���¼������
    For Each objFree In txtObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Text = ""
        End If
    Next
    
    For Each objFree In txtAge
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Text = ""
        End If
    Next
    
    For Each objFree In cbxAge
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.ListIndex = 0
        End If
    Next
    
    For Each objFree In lstObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Call objFree.Clear
        End If
    Next
    
    For Each objFree In cbxObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Text = ""
        End If
    Next
    
    For Each objFree In dtpObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Value = CurServerDate
        End If
    Next
    
    For Each objFree In cbxDateUnit
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.ListIndex = 0
        End If
    Next
    
    
    For Each objFree In chkObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Value = 0
        End If
    Next
    
    mblnIsLoading = False
    
    '�������õ�¼������
'    Call LoadInputData
        
Exit Sub
errHandle:
    mblnIsLoading = False
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Public Function UpdateFindCondition() As TSchemeItem
'���ò��һ�������
On Error GoTo errHandle
    Dim i As Long
    Dim inputTag As TInputTag
    
    For i = 1 To UBound(maryInputTag)
        inputTag = maryInputTag(i)
        
        If inputTag.ParName = "ϵͳ.��ʼ����" Then
            mobjSchemeItem.Startdate = inputTag.InputControl.Value
            mobjFilterValue.ParData("ϵͳ.��ʼ����") = mobjSchemeItem.Startdate
        ElseIf inputTag.ParName = "ϵͳ.��������" Then
            mobjSchemeItem.EndDate = inputTag.InputControl.Value
            mobjFilterValue.ParData("ϵͳ.��������") = mobjSchemeItem.EndDate
        Else
            If inputTag.ControlType = ctChk Then
                Call mobjFilterValue.UpdateParValue("@" & inputTag.ParName, _
                                            GetControlValue(inputTag.InputControl, inputTag))
            Else
                '�ж�ƥ�䷽ʽ
                Call mobjFilterValue.UpdateParValue(inputTag.ParName, _
                                            GetControlValue(inputTag.InputControl, inputTag, inputTag.LikeWay))
            End If
        End If
    Next i
    
    UpdateFindCondition = mobjSchemeItem
    
Exit Function
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Function

Public Sub Restore(objFilterValue As clsSqlFilterValue)
'�ָ���ʼ�Ĳ��������������¼�������
    Call ReadUserInputConfig(objFilterValue)
End Sub

Private Sub cmdSure_Click()
    Call UpdateFindCondition
    
    Me.Hide
    
    mblnIsOK = True
End Sub

Private Sub dtpObj_Change(Index As Integer)
'���ڿ�����ֵ���û��ı����Ҫ��������ݼ���
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index))
    
    If sdrRange.Visible Then
        If Index = sdrRange.Tag Or Index = sdrRange.Tag + 1 Then
            sdrRange.Value = dtpObj(sdrRange.Tag + 1).Value - dtpObj(sdrRange.Tag).Value
        End If
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub



Public Sub LoadFace()
On Error GoTo errHandle
      
    mblnIsLoading = True
    
    Set mobjSqlParse = New clsSqlParse
    
    Call ConfigBaseWindow
    Call ConfigTitleDisplay
    Call ConfigInputControl
    
    Call UpdateWindowSize
    
    'Call AutoHide
    
    Call sctExecute.AddObject("Me", Me, True)
    
    Call LoadBaseInputData
    
    '����filtervalue���ý�������
    Call ReadUserInputConfig(mobjFilterValue)
    
    mblnIsLoading = False
Exit Sub
errHandle:
    mblnIsLoading = False
    MsgBox "��ѯ���˴��ڼ���ʧ��:" & Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub Form_Load()
    If mblnIsEmbed Then Exit Sub
    
    Call LoadFace
End Sub


Private Sub AutoHide()
'�ؼ��Զ����ش���
On Error GoTo errHandle
    Dim obj As Object
    
    mblnIsMoreInput = False
    
    If mblnIsEmbed = False Then Exit Sub
    
    For Each obj In Me.Controls
        Call ControlVisible(obj)
    Next
    
Exit Sub
errHandle:
    Err.Clear
End Sub

Private Sub ControlVisible(obj As Object)
On Error Resume Next
    If obj.Name <> "cbxWhere" Then
        If Val(obj.Index) <= 0 Then Exit Sub
    End If
    
    
    obj.Visible = IIf(obj.Top + obj.Height > Me.ScaleHeight, False, True)
    
    If obj.Name = "cbxWhere" Then
        obj.Visible = obj.Visible And obj.ListCount > 0
    End If
    
    
    '����пؼ������أ����ʾ���и���¼��ؼ�û�б���ʾ����
    If obj.Visible = False Then mblnIsMoreInput = True
    
Err.Clear
End Sub

Private Sub ReadUserInputConfig(objFilterValue As clsSqlFilterValue)
On Error GoTo errHandle
    Dim i As Long
    
    If Not objFilterValue Is Nothing Then
        For i = 1 To objFilterValue.Count
            Call UpdateInputData(objFilterValue.Item(i).Name, objFilterValue.Item(i).Value)
        Next i
    End If
Exit Sub
errHandle:
    Debug.Print "ReadUserConfig Err:" & Err.Description
End Sub

Public Function GetFromData(ByVal strSql As String) As ADODB.Recordset
'��ȡ��Դ����
On Error GoTo errHandle
    Dim strQuerySql As String
    
    Set GetFromData = Nothing
    
    Call mobjSqlParse.init(strSql)
    
    strQuerySql = mobjSqlParse.GetQuerySql(False)
    
    Set GetFromData = ExecuteCore(strQuerySql, "��ȡ��������", mobjSqlParse.ParValues)
Exit Function
errHandle:
    Err.Raise -1, "frmQueryFilter.GetFromData", "[GetFromData]�������>>" & vbCrLf & "  ��ѯ���Ϊ��" & strSql & vbCrLf & Err.Description
    Resume
End Function

Private Sub LoadBaseInputData()
'���ؿ�ѡ¼������
'¼�����������������20
    Dim i As Long
    
    Dim inputTag As TInputTag
    Dim inputLen As Long

    
    
    inputLen = UBound(maryInputTag)
    For i = 1 To inputLen
        inputTag = maryInputTag(i)
        
        Call ConfigControlValue(inputTag, True)
    Next i
End Sub

Private Sub ConfigControlValue(ByRef inputTag As TInputTag, ByVal blnIsSetDefault As Boolean)
'����������Դ���ÿؼ�¼��ֵ
    Dim i As Long
    Dim lngInputType As Long
    Dim strTextDataSource() As String
    Dim rsSqlDataSource As ADODB.Recordset
    Dim strDefaultValue As String
    Dim objInputControl As Object
    Dim strDataItem As String
    
    lngInputType = inputTag.ControlType
    
    If inputTag.FromType = dbftText Then
        strTextDataSource = Split(inputTag.DataFrom, ";")
    Else
        Set rsSqlDataSource = GetFromData(inputTag.DataFrom)
    End If
    
    strDefaultValue = RunScripting(sctExecute, inputTag.Default)
    
    Set objInputControl = inputTag.InputControl
    
    Select Case lngInputType

        Case 0
            '��ȡ�ı�����ʾ������
            If inputTag.FromType = dbftText Then
                Call SetControlValue(objInputControl, inputTag.ControlType, inputTag.DataFrom)
            Else
                Call SetControlValue(objInputControl, inputTag.ControlType, rsSqlDataSource(0).Value)
            End If

            If strDefaultValue <> "" Then
                Call SetControlValue(objInputControl, inputTag.ControlType, strDefaultValue)
            End If
        Case 1, 2, 3, 9
            '��ȡ���ڿ���ʾ������
            If inputTag.FromType = dbftText Then
                Call SetControlValue(objInputControl, inputTag.ControlType, Now)
            Else
                Call SetControlValue(objInputControl, inputTag.ControlType, rsSqlDataSource(0).Value)
            End If

            If strDefaultValue <> "" Then
                Call SetControlValue(objInputControl, inputTag.ControlType, strDefaultValue)
            End If
        Case 4
            '��ȡ��������ʾ������
            objInputControl.Clear
            
            If inputTag.FromType = dbftText Then
                
                objInputControl.AddItem ""
                
                For i = 0 To UBound(strTextDataSource)
                    If i >= 100 Then Exit For
                    
                    strDataItem = strTextDataSource(i)
                    
                    If Trim$(strDataItem) <> "" Then
                        objInputControl.AddItem ParseInputValue(strDataItem, False)
'                        objInputControl.ItemData(objInputControl.ListCount - 1) = Val(ParseInputValue(strDataItem, True))
                        maryItemData(inputTag.Index, objInputControl.ListCount - 1) = ParseInputValue(strDataItem, True)
                    End If
                Next i
            Else
                objInputControl.AddItem ""
                
                i = 0
                Do While Not rsSqlDataSource.EOF
                    If i >= 100 Then Exit Do
                    i = i + 1
                    
                    strDataItem = rsSqlDataSource(0).Value
                    
                    If Trim$(strDataItem) <> "" Then
                        objInputControl.AddItem ParseInputValue(strDataItem, False)
'                        objInputControl.ItemData(objInputControl.ListCount - 1) = Val(ParseInputValue(strDataItem, True))
                        maryItemData(inputTag.Index, objInputControl.ListCount - 1) = ParseInputValue(strDataItem, True)
                    End If
                    
                    rsSqlDataSource.MoveNext
                Loop
            End If

            If strDefaultValue <> "" Then
                Call SetControlValue(objInputControl, inputTag.ControlType, strDefaultValue)
            Else
                If objInputControl.ListCount > 0 Then
                    objInputControl.ListIndex = 0
                Else
                    objInputControl.Text = ""
                End If
            End If
        Case 5
            '��ȡ�ɶ�ѡ�б����ʾ������
            objInputControl.Clear
            
            If inputTag.FromType = dbftText Then
                For i = 0 To UBound(strTextDataSource)
                    If i >= 100 Then Exit For
                    i = i + 1
                    
                    strDataItem = strTextDataSource(i)
                    
                    If Trim$(strDataItem) <> "" Then
                        objInputControl.AddItem ParseInputValue(strDataItem, False)
'                        objInputControl.ItemData(objInputControl.ListCount - 1) = Val(ParseInputValue(strDataItem, True))
                        maryItemData(inputTag.Index, objInputControl.ListCount - 1) = ParseInputValue(strDataItem, True)
                    End If
                    
                    If InStr(strDefaultValue, strDataItem) > 0 Then
                        objInputControl.Selected(objInputControl.ListCount - 1) = True
                    End If
                Next i
            Else
                i = 0
                Do While Not rsSqlDataSource.EOF
                    If i >= 100 Then Exit Do
                    i = i + 1
                    
                    strDataItem = rsSqlDataSource(0).Value
                    
                    If Trim$(strDataItem) <> "" Then
                        objInputControl.AddItem ParseInputValue(strDataItem, False)
'                        objInputControl.ItemData(objInputControl.ListCount - 1) = Val(ParseInputValue(strDataItem, True))
                        maryItemData(inputTag.Index, objInputControl.ListCount - 1) = ParseInputValue(strDataItem, True)
                    End If

                    If InStr(strDefaultValue, rsSqlDataSource(0).Value) > 0 Then
                        objInputControl.Selected(objInputControl.ListCount - 1) = True
                    End If

                    rsSqlDataSource.MoveNext
                Loop
            End If
        Case 8
            
    End Select
End Sub

Private Function ParseInputValue(ByVal strSourceValue As String, ByVal blnIsItemData As Boolean) As String
On Error GoTo errHandle
    Dim lngSplitIndex As Long
    
    ParseInputValue = strSourceValue
    
    If InStr(Trim$(strSourceValue), "@") = 1 Then
        If blnIsItemData = False Then
            ParseInputValue = Mid(strSourceValue, 2, 255)
        Else
            ParseInputValue = ""
        End If
        Exit Function
    End If
    
    lngSplitIndex = InStr(strSourceValue, "-")
    
    If blnIsItemData Then
        If lngSplitIndex <= 0 Then
            ParseInputValue = ""
        Else
            ParseInputValue = Mid(strSourceValue, 1, lngSplitIndex - 1)
        End If
    Else
        If lngSplitIndex <= 0 Then
            ParseInputValue = strSourceValue
        Else
            ParseInputValue = Mid(strSourceValue, lngSplitIndex + 1, 255)
        End If
    End If
Exit Function
errHandle:
    ParseInputValue = ""
End Function

Private Sub SetControlValue(objInputControl As Object, ByVal lngInputType As Long, ByVal strValue As Variant)
'�Կؼ����ı�����value���Ը�ֵ
On Error Resume Next
    Dim i As Long
    
    Select Case lngInputType
        Case ctText '�ı���
            objInputControl.Text = strValue
        Case ctDate, ctDateTime, ctTime, ctFastDate '���ڿ�
            objInputControl.Value = strValue
        Case ctCombobox '������
            objInputControl.Text = strValue
        Case ctList '�б��
            For i = 0 To objInputControl.ListCount - 1
                If objInputControl.List(i) = strValue Then
                    objInputControl.Selected(i) = True
                End If
            Next i
        Case ctChk  '��ѡ��
            If CBool(strValue) <> False Then objInputControl.Value = 1
    End Select
End Sub

Private Sub ClearControlValue(objInputControl As Object, ByVal lngInputType As Long)
    Select Case lngInputType
        Case ctText '�ı���
            objInputControl.Text = ""
        Case ctCombobox '������
            Call objInputControl.Clear
        Case ctList '�б��
            Call objInputControl.Clear
    End Select
End Sub

Private Sub UpdateWindowSize()
    If Not mobjLastControl Is Nothing Then
        framButton.Top = mobjLastControl.Top + mobjLastControl.Height + 120 + 15
        Me.Height = framButton.Top + framButton.Height + 400 - 15
        
        cmdClear.Top = framButton.Top + 240 * mdblZoomRate
        cmdCancel.Top = framButton.Top + 240 * mdblZoomRate
        cmdSure.Top = framButton.Top + 240 * mdblZoomRate
        
        labError.Visible = False
    Else
        labError.Visible = True
    End If
End Sub

Private Sub ConfigBaseWindow()
    If mdblFontSize > 0 Then
        Call SetFontSize(mdblFontSize)
        Me.FontSize = mdblFontSize
    End If
    
    mdblZoomRate = 1
    If mdblFontSize > 10 Then mdblZoomRate = 1 + (mdblFontSize / 2 - 5) / 10
    
    Me.Width = 5724 * mdblZoomRate
    
    If mblnIsEmbed Then
        shpBack.Visible = False
        labMemo.Visible = False
        imgQuery.Visible = False
        
        framButton.Visible = False
        
        cmdClear.Visible = False
        cmdCancel.Visible = False
        cmdSure.Visible = False
    Else
        shpBack.Width = 5675 * mdblZoomRate
        labMemo.Width = 4656 * mdblZoomRate
            
        framButton.Left = -45
        framButton.Width = Me.ScaleWidth + 90
        framButton.Height = 795 * mdblZoomRate
        
        
        cmdClear.Width = 1300 * mdblZoomRate
        cmdClear.Height = 375 * mdblZoomRate
        
        
        cmdCancel.Width = 1300 * mdblZoomRate
        cmdCancel.Height = 375 * mdblZoomRate
        cmdCancel.Left = Me.Width - cmdCancel.Width - 240
        
        cmdSure.Width = 1300 * mdblZoomRate
        cmdSure.Height = 375 * mdblZoomRate
        cmdSure.Left = cmdCancel.Left - cmdSure.Width - 240
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 5 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
On Error GoTo errHandle
    Call AutoHide
Exit Sub
errHandle:
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjSqlParse = Nothing
End Sub

Private Sub lstObj_Click(Index As Integer)
'��ѡ������ֵ���û��ı����Ҫ��������ݼ���
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index))
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub mobjSqlParse_OnGetParameterValue(ByVal strParName As String, Value As Variant)
    '��ȡ����
    Dim i As Long
    Dim inputTag As TInputTag
    
    For i = 1 To UBound(maryInputTag)
        inputTag = maryInputTag(i)
        If inputTag.ParName = strParName Then
            Value = GetControlValue(inputTag.InputControl, inputTag)
        End If
    Next i
End Sub

Private Function GetControlValue(objInputControl As Object, ByRef inputTag As TInputTag, Optional ByVal lngLikeWay As TLikeWay = lwNormal) As Variant
    Dim i As Long
    
    Select Case inputTag.ControlType
        Case ctText  'textbox�ı�
            GetControlValue = objInputControl.Text
        Case ctDate  'dtpicker����
            GetControlValue = CDate(Format(objInputControl.Value, "yyyy-MM-dd"))
        Case ctTime  'dtpickerʱ��
            GetControlValue = CDate(Format(objInputControl.Value, "HH:mm"))
        Case ctDateTime  'dtpicker����ʱ��
            GetControlValue = CDate(Format(objInputControl.Value, "yyyy-MM-dd HH:mm"))
        Case ctCombobox  'combobox����
            GetControlValue = objInputControl.Text
            
            If Trim(objInputControl.Text) = "" Then Exit Function
            
            If objInputControl.ListIndex >= 0 Then
'                If objInputControl.ItemData(objInputControl.ListIndex) <> 0 Then
'                    GetControlValue = objInputControl.ItemData(objInputControl.ListIndex)
'                End If
                If maryItemData(inputTag.Index, objInputControl.ListIndex) <> "" Then
                    GetControlValue = maryItemData(inputTag.Index, objInputControl.ListIndex)
                End If
            Else
                For i = 0 To objInputControl.ListCount - 1
                    If objInputControl.List(i) = objInputControl.Text Then
                        If maryItemData(inputTag.Index, i) <> "" Then
                            GetControlValue = maryItemData(inputTag.Index, i)
                        End If
                        
                        Exit Function
                    End If
                Next i
            End If
        Case ctList  'list�б�
            For i = 0 To objInputControl.ListCount - 1
                If objInputControl.Selected(i) Then
                    If GetControlValue <> "" Then GetControlValue = GetControlValue & ","
'                    If objInputControl.ItemData(i) <> 0 Then
'                        GetControlValue = GetControlValue & objInputControl.ItemData(i)
'                    Else
'                        GetControlValue = GetControlValue & objInputControl.List(i)
'                    End If
                    If maryItemData(inputTag.Index, i) <> "" Then
                        GetControlValue = GetControlValue & maryItemData(inputTag.Index, i)
                    Else
                        GetControlValue = GetControlValue & objInputControl.List(i)
                    End If
                End If
            Next i
        Case ctChk  'checkbox��ѡ��
            GetControlValue = IIf(objInputControl.Value <> 0, True, False)
            
        Case ctAgeCbx  '������Ͽ�
            If Trim(objInputControl.Text) = "" Then Exit Function
            
            GetControlValue = GetAgeDays(objInputControl.Text, cbxAge(objInputControl.Index).Text)
        Case ctMutxCbx  '��n�����������
            If objInputControl.Text = inputTag.ParName Then
                GetControlValue = txtObj(objInputControl.Tag).Text
            End If
        Case ctFastDate  '���ڿ�ѡ���
            GetControlValue = CDate(Format(objInputControl.Value, "yyyy-MM-dd"))
        
    End Select
    
    If lngLikeWay <> lwNormal Then
        If IsEmpty(GetControlValue) Or IsNull(GetControlValue) Or GetControlValue = "" Then Exit Function
        
        Select Case lngLikeWay
            Case lwLeft
                GetControlValue = "%" & GetControlValue
            Case lwRight
                GetControlValue = GetControlValue & "%"
            Case lwAll
                GetControlValue = "%" & GetControlValue & "%"
        End Select
    End If
End Function

Private Function GetAgeDays(ByVal strAge As String, ByVal strUnit As String) As Long
'ת��Ϊ��������
    Select Case strUnit
        Case "S-��"
            GetAgeDays = Val(strAge) * 365
        Case "Y-��"
            GetAgeDays = Val(strAge) * 30
        Case "Z-��"
            GetAgeDays = Val(strAge) * 7
        Case "T-��"
            GetAgeDays = Val(strAge) * 1
    End Select
End Function

Private Sub sdrRange_Scroll()
On Error GoTo errHandle
    dtpObj(Val(sdrRange.Tag)).Value = dtpObj(sdrRange.Tag + 1).Value - sdrRange.Value
    
Exit Sub
errHandle:
    Debug.Print "ERR>>sdrRange_Scroll:" & Err.Description
End Sub

Private Sub txtAge_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtObj_Change(Index As Integer)
'�ı�������ֵ���û��ı����Ҫ��������ݼ���
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index))
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub
