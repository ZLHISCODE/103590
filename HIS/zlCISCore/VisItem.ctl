VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl VisItem 
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   ScaleHeight     =   435
   ScaleWidth      =   6375
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "VisItem.ctx":0000
         Left            =   2640
         List            =   "VisItem.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1250
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   2400
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         OrigLeft        =   3000
         OrigTop         =   1560
         OrigRight       =   3240
         OrigBottom      =   1830
         Enabled         =   -1  'True
      End
      Begin VB.Line line_Under 
         BorderColor     =   &H80000004&
         X1              =   1680
         X2              =   2160
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Left            =   600
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   660
      End
   End
End
Attribute VB_Name = "VisItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const Margin As Integer = 20
Private iViewMethod As Integer, sTitle As String, iType As Integer, iMaxLength As Integer, iDecLength As Integer
Private ItemID As String
Private aItemValues() As String
Private sDefaultValue As String
Private bEnabled As Boolean
Private bAllowEdit As Boolean
Private sUnit As String
Private bMask As Boolean
Private sExchangedFld As String
Private iTextMethod As Integer, sNullString As String, sFormat As String

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private bTxtChanged As Boolean

'BEGIN BY CFR 2005-06-10
Private mobjParentObject As Object

Public Property Set ParentObject(vData As Object)
    Set mobjParentObject = vData
End Property

Public Property Get ParentObject() As Object
    Set ParentObject = mobjParentObject
End Property

Private Property Let Modified(vData As Boolean)
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    mobjParentObject.Modified = vData
    
End Property

Private Property Get Modified() As Boolean
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    Modified = mobjParentObject.Modified
    
End Property

'END BY CFR 2005-06-10

Public Sub Init(ByVal ItemTitle As String, ByVal Unit As String, ByVal ViewMethod As Integer, Optional ByVal ItemType As Integer = 0, Optional ByVal MaxLength As String = "0", Optional ByVal DecLength As String = "0", Optional ByVal ItemValues As String = "", Optional ByVal DefaultValue As String = "", Optional ByVal ID As String = "", Optional ByVal ExchangedFld As String = "", Optional ByVal TextMethod As Integer = 1, Optional ByVal NullString As String = "", Optional ByVal strFormat As String)
    Dim iValueNum As Long, i As Long
    Dim LeftMargin As Long, Seq As Long
    On Error Resume Next
    iViewMethod = ViewMethod
    sTitle = Trim(ItemTitle)
    sUnit = Unit
    iType = ItemType
    iMaxLength = CInt(MaxLength)
    iDecLength = CInt(DecLength)
    aItemValues = Split(ItemValues, ";")
    sDefaultValue = Trim(DefaultValue)
    ItemID = ID
    sExchangedFld = ExchangedFld
    iTextMethod = TextMethod: sNullString = NullString: sFormat = strFormat
    
    bTxtChanged = False
    
    Select Case True '表示法
        Case iViewMethod = 2 And iType <> 3 '下拉
            With Combo1
                .Clear
                .AddItem ""
                
                iValueNum = -1
                iValueNum = UBound(aItemValues)
                For i = 0 To iValueNum
                    .AddItem aItemValues(i)
                Next
            End With
        Case ViewMethod = 3 Or ItemType = 3 '复选
            iValueNum = -1
            iValueNum = UBound(aItemValues)
            If ItemType <> 3 And iValueNum > -1 Then
                For i = 0 To iValueNum
                    Load Check1(Check1.Count)
                Next
            End If
        Case ViewMethod = 4 And ItemType <> 3 '单选
            iValueNum = -1
            iValueNum = UBound(aItemValues)
            If iValueNum > -1 Then
                For i = 0 To iValueNum
                    Load Option1(Option1.Count)
                Next
            End If
    End Select
    ShowItem
    
    SetValue sDefaultValue
    SetControlEnabled
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "背景色"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";外观"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    Dim i As Long
    UserControl.BackColor = vNewValue
    picMain.BackColor = vNewValue
    For i = 0 To Check1.Count - 1
        Check1(i).BackColor = vNewValue
    Next
    For i = 0 To Option1.Count - 1
        Option1(i).BackColor = vNewValue
    Next
End Property

Private Sub Check1_Click(Index As Integer)
    Dim iNum As Long, i As Long
    On Error Resume Next
    If iType = 3 Then
        With Check1(Index)
            sDefaultValue = IIf(.Value = 1, "是", "否")
        End With
    Else
        Select Case True '表示法
            Case (iViewMethod = 0 And iType <> 3) Or (iViewMethod = 1 And iType = 1) '文本
                If Check1(Index).Value = 1 Then
                    Text1.Enabled = True
                    Text1.SetFocus
                Else
                    Text1.Text = ""
                    Text1.Enabled = False
                    Check1(0).SetFocus
                    
                    sDefaultValue = ""
                End If
            Case iViewMethod = 1 And (iType = 0 Or iType = 2) '范围
                If Check1(Index).Value = 1 Then
                    Text1.Enabled = True
                    Text1.SetFocus
                Else
                    Text1.Text = ""
                    Text1.Enabled = False
                    Check1(0).SetFocus
                    
                    sDefaultValue = ""
                End If
            Case iViewMethod = 2 And iType <> 3 '下拉
                If Check1(Index).Value = 1 Then
                    Combo1.Enabled = True
                    Combo1.SetFocus
                Else
                    Combo1.ListIndex = 0
                    Combo1.Enabled = False
                    Check1(0).SetFocus
                    
                    sDefaultValue = ""
                End If
            Case Else
                sDefaultValue = ""
                
                iNum = Check1.Count - 1
                For i = 1 To iNum
                    If Check1(i).Value = 1 Then sDefaultValue = sDefaultValue + ";" + aItemValues(i - 1)
                Next
                If Len(sDefaultValue) > 0 Then sDefaultValue = Mid(sDefaultValue, 2)
        End Select
    End If
    
    Modified = True
    
End Sub
'
'Private Sub Check1_GotFocus(Index As Integer)
'    Select Case True '表示法
'        Case (iViewMethod = 0 And iType <> 3) Or (iViewMethod = 1 And iType = 1) '文本
'            If Check1(Index).Value = 1 Then
'                Text1.Enabled = True
'                Text1.SetFocus
'            End If
'        Case iViewMethod = 1 And (iType = 0 Or iType = 2) '范围
'            If Check1(Index).Value = 1 Then
'                Text1.Enabled = True
'                Text1.SetFocus
'            End If
'        Case iViewMethod = 2 And iType <> 3 '下拉
'            If Check1(Index).Value = 1 Then
'                Combo1.Enabled = True
'                Combo1.SetFocus
'            End If
'    End Select
'End Sub

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If bMask And Check1(0).Value = 1 Then
        zlCommFun.PressKey vbKeyTab
    Else
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
'
'Private Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If bMask Then
'        Check1(Index).Value = Check1(Index) Xor 1
'    End If
'End Sub

Private Sub Combo1_Click()
    sDefaultValue = Combo1.Text
    
    Modified = True
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Option1_Click(Index As Integer)
    sDefaultValue = IIf(Option1(Index), aItemValues(Index - 1), "")
    
    Modified = True
    
End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        If Index >= Option1.Count - 1 Then
'            RaiseEvent KeyPress(KeyAscii)
'        Else
'            Option1(Index + 1).SetFocus
'        End If
'    Else
        RaiseEvent KeyPress(KeyAscii)
'    End If
End Sub

Private Sub picMain_GotFocus()
    RaiseEvent KeyDown(vbKeyReturn, 0)
End Sub

Private Sub Text1_Change()
    Modified = True
End Sub

Private Sub Text1_GotFocus()
    With Text1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If iViewMethod = 1 And iType = 0 Then
        If KeyCode = vbKeyUp Then UpDown1_UpClick
        If KeyCode = vbKeyDown Then UpDown1_DownClick
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If iViewMethod < 2 And iType <> 3 Then
        If Not ifEditKey(KeyAscii, False) And LenB(StrConv(Text1.Text, vbFromUnicode)) >= iMaxLength And Text1.SelLength = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
        
        Select Case iType
            Case 0
                If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or ifEditKey(KeyAscii, False)) Then
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
        
        bTxtChanged = True
    End If
    If KeyAscii > 0 Then RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_LostFocus()
    Dim ValuesCount As Long
    
    If Not bTxtChanged Then Exit Sub
    If iViewMethod < 2 And iType <> 3 And Len(Trim(Text1)) > 0 Then
        On Error Resume Next
        ValuesCount = 0
        ValuesCount = UBound(aItemValues) + 1
        
        Select Case iType
            Case 0
                If Not IsNumeric(Text1) Then
                    MsgBox "必须输入数字！", vbExclamation + vbOKOnly, gstrSysName
                    Text1.SetFocus
                    Exit Sub
                End If
                If iMaxLength > 0 And Abs(CLng(Text1)) > 0 And Len(CStr(Abs(CLng(Text1)))) > CLng(iMaxLength) - CLng(iDecLength) Then
                    MsgBox "值超过定义的最大长度！", vbExclamation + vbOKOnly, gstrSysName
                    Text1.SetFocus
                    Exit Sub
                End If
                If ValuesCount > 0 Then
                    If Len(Trim(aItemValues(0))) > 0 Then
                        If CDbl(Text1) < CDbl(aItemValues(0)) Then
                            MsgBox "输入值不能小于" & aItemValues(0) & "！", vbExclamation + vbOKOnly, gstrSysName
                            Text1.SetFocus
                            Exit Sub
                        End If
                    End If
                    If ValuesCount > 1 Then
                        If Len(Trim(aItemValues(1))) > 0 Then
                            If CDbl(Text1) > CDbl(aItemValues(1)) Then
                                MsgBox "输入值不能大于" & aItemValues(1) & "！", vbExclamation + vbOKOnly, gstrSysName
                                Text1.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                If CInt(iDecLength) > 0 Then
                    Text1 = Format(Round(Text1, CLng(iDecLength)), "#." + String(CLng(iDecLength), "0"))
                Else
                    Text1 = Round(Text1, 0)
                End If
            Case 1
                If iMaxLength > 0 And Len(Trim(Text1)) > CLng(iMaxLength) Then
                    MsgBox "值超过定义的最大长度！", vbExclamation + vbOKOnly, gstrSysName
                    Text1.SetFocus
                    Exit Sub
                End If
                If ValuesCount > 0 Then
                    If Len(Trim(aItemValues(0))) > 0 Then
                        If Text1 < aItemValues(0) Then
                            MsgBox "输入值不能小于" & aItemValues(0) & "！", vbExclamation + vbOKOnly, gstrSysName
                            Text1.SetFocus
                            Exit Sub
                        End If
                    End If
                    If ValuesCount > 1 Then
                        If Len(Trim(aItemValues(1))) > 0 Then
                            If Text1 > aItemValues(1) Then
                                MsgBox "输入值不能大于" & aItemValues(1) & "！", vbExclamation + vbOKOnly, gstrSysName
                                Text1.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                If Len(sFormat) > 0 Then Text1 = Format(Text1, sFormat)
            Case 2
                If Not IsDate(Text1) Then
                    MsgBox "错误的日期格式！日期格式为" + IIf(Len(sFormat) = 0, "YYYY-MM-DD HH:MM:SS", sFormat), vbExclamation + vbOKOnly, gstrSysName
                    Text1.SetFocus
                    Exit Sub
                Else
                    Text1 = Format(Text1, IIf(Len(sFormat) = 0, "YYYY-MM-DD HH:MM:SS", sFormat))
                End If
                If ValuesCount > 0 Then
                    If Len(Trim(aItemValues(0))) > 0 Then
                        If CDate(Text1) < CDate(aItemValues(0)) Then
                            MsgBox "输入值不能小于" & aItemValues(0) & "！", vbExclamation + vbOKOnly, gstrSysName
                            Text1.SetFocus
                            Exit Sub
                        End If
                    End If
                    If ValuesCount > 1 Then
                        If Len(Trim(aItemValues(1))) > 0 Then
                            If CDate(Text1) > CDate(aItemValues(1)) Then
                                MsgBox "输入值不能大于" & aItemValues(1) & "！", vbExclamation + vbOKOnly, gstrSysName
                                Text1.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    sDefaultValue = Text1
    bTxtChanged = False
End Sub

Private Sub UpDown1_DownClick()
    Dim MinValue As String
    Dim MaxValue As String
    
    On Error Resume Next
    MinValue = ""
    MinValue = aItemValues(0)
    MaxValue = ""
    MaxValue = aItemValues(1)
    
    Text1 = CDbl(Text1) - 1
    If Len(MinValue) > 0 Then
        If CDbl(Text1) < CDbl(MinValue) Then Text1 = MinValue
    End If
    If Len(MaxValue) > 0 Then
        If CDbl(Text1) > CDbl(MaxValue) Then Text1 = MaxValue
    End If
    
    sDefaultValue = Text1
    bTxtChanged = True
End Sub

Private Sub UpDown1_UpClick()
    Dim MinValue As String
    Dim MaxValue As String
    
    On Error Resume Next
    MinValue = ""
    MinValue = aItemValues(0)
    MaxValue = ""
    MaxValue = aItemValues(1)
    
    Text1 = CDbl(Text1) + 1
    If Len(MinValue) > 0 Then
        If CDbl(Text1) < CDbl(MinValue) Then Text1 = MinValue
    End If
    If Len(MaxValue) > 0 Then
        If CDbl(Text1) > CDbl(MaxValue) Then Text1 = MaxValue
    End If
    
    sDefaultValue = Text1
    bTxtChanged = True
End Sub

Private Sub UserControl_GotFocus()
    RaiseEvent KeyDown(vbKeyReturn, 0)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim i As Long
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    For i = 0 To Check1.Count - 1
        Check1(i).BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    Next
    For i = 0 To Option1.Count - 1
        Option1(i).BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    Next

    UserControl.MousePointer = PropBag.ReadProperty("MousePointer") ', ccArrow)
    
    Set UserControl.Font = PropBag.ReadProperty("Font", UserControl.Font)
    SetFont
    
    bEnabled = PropBag.ReadProperty("Enabled", True)
    bAllowEdit = PropBag.ReadProperty("AllowEdit", False)
    
    SetControlEnabled
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With picMain
        .Left = 0: .Top = 0
        .Width = UserControl.Width: .Height = UserControl.Height
    End With
    Select Case True '表示法
        Case (iViewMethod = 0 And iType <> 3) Or (iViewMethod = 1 And iType = 1) '文本
            If UserControl.Width - Margin - lblUnit.Width - 100 - Text1.Left < 300 Then UserControl.Width = 300 + Margin + lblUnit.Width + 100 + Text1.Left
            With Text1
                .Width = UserControl.Width - Margin - lblUnit.Width - 100 - .Left
            End With
            line_Under.X2 = Text1.Width + line_Under.X1
            With lblUnit
                .Left = Text1.Left + Text1.Width + 100
            End With
            UserControl.Height = Text1.Height + 2 * Margin
        Case iViewMethod = 1 And (iType = 0 Or iType = 2) '范围
            If UserControl.Width - Margin - lblUnit.Width - 100 - UpDown1.Width - Text1.Left < 300 Then UserControl.Width = 300 + Margin + lblUnit.Width + 100 + UpDown1.Width + Text1.Left
            With Text1
                .Width = UserControl.Width - Margin - lblUnit.Width - 100 - UpDown1.Width - .Left
            End With
            line_Under.X2 = Text1.Width + line_Under.X1
            With UpDown1
                .Left = Text1.Left + Text1.Width
            End With
            With lblUnit
                .Left = UpDown1.Left + UpDown1.Width + 100
            End With
            UserControl.Height = Text1.Height + 2 * Margin
        Case iViewMethod = 2 And iType <> 3 '下拉
            If UserControl.Width - Margin - lblUnit.Width - 100 - Combo1.Left < 300 Then UserControl.Width = 300 + Margin + lblUnit.Width + 100 + Combo1.Left
            With Combo1
                .Width = UserControl.Width - Margin - lblUnit.Width - 100 - .Left
            End With
            With lblUnit
                .Left = Combo1.Left + Combo1.Width + 100
            End With
            UserControl.Height = Combo1.Height + 2 * Margin
        Case iViewMethod = 3 Or iType = 3 '复选
'            If UserControl.Width - Margin - Check1(Check1.Count - 1).Left < UserControl.TextWidth(Check1(Check1.Count - 1).Caption) + 300 Then UserControl.Width = UserControl.TextWidth(Check1(Check1.Count - 1).Caption) + 300 + Margin + Check1(Check1.Count - 1).Left
'            With Check1(Check1.Count - 1)
'                .Width = UserControl.Width - Margin - .Left
'                UserControl.Height = .Height + 2 * Margin
'            End With
            ShowItem
        Case iViewMethod = 4 And iType <> 3 '单选
'            If UserControl.Width - Margin - Option1(Option1.Count - 1).Left < UserControl.TextWidth(Option1(Option1.Count - 1).Caption) + 300 Then UserControl.Width = UserControl.TextWidth(Option1(Option1.Count - 1).Caption) + 300 + Margin + Option1(Option1.Count - 1).Left
'            With Option1(Option1.Count - 1)
'                .Width = UserControl.Width - Margin - .Left
'                UserControl.Height = .Height + 2 * Margin
'            End With
            ShowItem
    End Select
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    
    Set mobjParentObject = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", UserControl.BackColor, vbWindowBackground
    PropBag.WriteProperty "MousePointer", UserControl.MousePointer ', ccArrow
    PropBag.WriteProperty "Font", UserControl.Font
    PropBag.WriteProperty "Enabled", bEnabled, True
    PropBag.WriteProperty "AllowEdit", bAllowEdit, False
End Sub

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "鼠标指针"
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";外观"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal vNewValue As MousePointerConstants)
    UserControl.MousePointer = vNewValue
End Property

Public Property Get Title() As String
Attribute Title.VB_Description = "标题"
    Title = sTitle
End Property

Public Property Let Title(ByVal vNewValue As String)
    Dim iValueNum As Long
    
    sTitle = Trim(vNewValue)
    
    ShowItem
End Property

Public Property Get ValuesCount() As Long
    On Error Resume Next
    ValuesCount = 0
    ValuesCount = UBound(aItemValues) + 1
End Property

Public Property Get Values(ByVal Index As Long) As String
    On Error Resume Next
    Values = vbNullString
    Values = aItemValues(Index)
End Property

Public Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    bEnabled = vNewValue
    SetControlEnabled
End Property

Public Property Get Value() As String
    Value = sDefaultValue
End Property

Public Property Let Value(ByVal vNewValue As String)
    sDefaultValue = vNewValue
    
    SetValue vNewValue
End Property

Public Property Get Method() As Integer
    Method = iViewMethod
End Property

Public Property Get ItemType() As Integer
    ItemType = iType
End Property

Public Property Get MaxLength() As Integer
    MaxLength = iMaxLength
End Property

Public Property Get DecLength() As Integer
    DecLength = iDecLength
End Property

Public Property Get ID() As String
    ID = ItemID
End Property

Private Sub SetValue(ByVal vNewValue As String)
    Dim i As Long, iNum As Long
    
    Dim blnSave As Boolean
    
    blnSave = Modified
    
    On Error Resume Next
    Select Case True '表示法
        Case (iViewMethod = 0 And iType <> 3) Or (iViewMethod = 1 And iType = 1) '文本
            With Text1
                .Text = vNewValue
            End With
            If Len(Trim(vNewValue)) > 0 Then Check1(0).Value = 1
        Case iViewMethod = 1 And (iType = 0 Or iType = 2) '范围
            With Text1
                .Text = vNewValue
            End With
            If Len(Trim(vNewValue)) > 0 Then Check1(0).Value = 1
        Case iViewMethod = 2 And iType <> 3 '下拉
            With Combo1
                If Len(Trim(vNewValue)) = 0 Then
                    .ListIndex = 0
                Else
                    .Text = vNewValue
                End If
            End With
            If Len(Trim(vNewValue)) > 0 Then Check1(0).Value = 1
        Case iViewMethod = 3 Or iType = 3 '复选
            If iType = 3 Then
                With Check1(Check1.Count - 1)
                    .Value = IIf(vNewValue = "是" Or vNewValue = "1", 1, 0)
                End With
            Else
                iNum = -1
                iNum = UBound(aItemValues)
                For i = 0 To iNum
                    If InStr(vNewValue, aItemValues(i)) > 0 Then
                        Check1(i + 1).Value = 1
                    Else
                        Check1(i + 1).Value = 0
                    End If
                Next
            End If
        Case iViewMethod = 4 And iType <> 3 '单选
            iNum = -1
            iNum = UBound(aItemValues)
            For i = 0 To iNum
                If aItemValues(i) = vNewValue Then
                    Option1(i + 1).Value = True
                Else
                    Option1(i + 1).Value = False
                End If
            Next
    End Select
    
    Modified = blnSave
    
End Sub

Public Property Get AllowEdit() As Boolean
Attribute AllowEdit.VB_Description = "是否允许编辑"
Attribute AllowEdit.VB_ProcData.VB_Invoke_Property = ";行为"
    AllowEdit = bAllowEdit
End Property

Public Property Let AllowEdit(ByVal vNewValue As Boolean)
    bAllowEdit = vNewValue
    SetControlEnabled
End Property

Private Sub SetControlEnabled()
    Dim bCanEdit As Boolean
    
    bCanEdit = IIf(bAllowEdit And bEnabled, True, False)
    
    picMain.Enabled = bCanEdit
End Sub

'判断是否为编辑键
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Or KeyAscii = vbKeyShift Or KeyAscii = vbKeyControl Or KeyAscii = vbKeyMenu Or _
      KeyAscii = vbKeyCapital Or KeyAscii = vbKeyPageUp Or KeyAscii = vbKeyPageDown Or _
      KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Or KeyAscii = vbKeyNumlock Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "显示字体"
Attribute Font.VB_ProcData.VB_Invoke_Property = ";外观"
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal vNewValue As StdFont)
    Set UserControl.Font = vNewValue
    
    SetFont
End Property

Private Sub SetFont()
    Dim tmpCtrl As Control
    
    On Error Resume Next
    For Each tmpCtrl In UserControl.Controls
        Set tmpCtrl.Font = UserControl.Font
    Next
End Sub

Public Property Get ExchangeField() As String
    ExchangeField = sExchangedFld
End Property
Private Sub ShowItem()
    Dim iValueNum As Long, i As Long
    Dim LeftMargin As Long, TopMargin As Long, Seq As Long
    Dim iMaxWidth As Long, iMaxHeight As Long
    Dim LeftStartMargin As Long
    
    Select Case True '表示法
        Case (iViewMethod = 0 And iType <> 3) Or (iViewMethod = 1 And iType = 1) '文本
            Seq = 0
            If bMask Then
                With Check1(0)
                    .Left = Margin
                    .Top = (Text1.Height - .Height) / 2 + Margin
                    .Caption = sTitle
                    .Width = UserControl.TextWidth(.Caption) + 300
                    
                    If Len(sDefaultValue) > 0 Then
                        .Value = 1
                        Text1.Enabled = True
                    Else
                        .Value = 0
                        Text1.Enabled = False
                    End If
                    
                    .TabIndex = 0: Seq = 1
                    .Visible = True
                    lblTitle.Visible = False
                    
                    LeftStartMargin = .Left + .Width + 100
                End With
            Else
                With lblTitle
                    .Left = Margin: .Top = (Text1.Height - .Height) / 2 + Margin
                    .Caption = sTitle
                    .Visible = True
                    Check1(0).Visible = False
                
                    LeftStartMargin = .Left + .Width + 100
                End With
            End If
            With Text1
                .Left = LeftStartMargin
                .Top = Margin
                .TabIndex = Seq
                .Visible = True
            End With
            With line_Under
                .X1 = Text1.Left
                .Y1 = Text1.Top + Text1.Height: .Y2 = .Y1
                .Visible = True
            End With
            With lblUnit
                .Left = Text1.Left + Text1.Width + 100: .Top = (Text1.Height - .Height) / 2 + Margin
                .Caption = sUnit
                .Visible = True
            End With
            UserControl.Width = lblUnit.Left + lblUnit.Width + Margin
            UserControl.Height = line_Under.Y1 + Margin
        Case iViewMethod = 1 And (iType = 0 Or iType = 2) '范围
            Seq = 0
            If bMask Then
                With Check1(0)
                    .Left = Margin
                    .Top = (Text1.Height - .Height) / 2 + Margin
                    .Caption = sTitle
                    .Width = UserControl.TextWidth(.Caption) + 300
                    
                    If Len(sDefaultValue) > 0 Then
                        .Value = 1
                        Text1.Enabled = True
                    Else
                        .Value = 0
                        Text1.Enabled = False
                    End If
                    
                    .TabIndex = 0: Seq = 1
                    .Visible = True
                    lblTitle.Visible = False
                    
                    LeftStartMargin = .Left + .Width + 100
                End With
            Else
                With lblTitle
                    .Left = Margin: .Top = (Text1.Height - .Height) / 2 + Margin
                    .Caption = sTitle
                    .Visible = True
                    Check1(0).Visible = False
                    
                    LeftStartMargin = .Left + .Width + 100
                End With
            End If
            With Text1
                .Left = LeftStartMargin
                .Top = Margin
                .TabIndex = Seq
                .Visible = True
            End With
            With line_Under
                .X1 = Text1.Left
                .Y1 = Text1.Top + Text1.Height: .Y2 = .Y1
                .Visible = True
            End With
            With UpDown1
                .Left = Text1.Left + Text1.Width: .Top = (Text1.Height - .Height) / 2 + Margin
                .Visible = True
            End With
            With lblUnit
                .Left = UpDown1.Left + UpDown1.Width + 100: .Top = (Text1.Height - .Height) / 2 + Margin
                .Caption = sUnit
                .Visible = True
            End With
            UserControl.Width = lblUnit.Left + lblUnit.Width + Margin
            UserControl.Height = line_Under.Y1 + Margin
        Case iViewMethod = 2 And iType <> 3 '下拉
            Seq = 0
            If bMask Then
                With Check1(0)
                    .Left = Margin
                    .Top = (Combo1.Height - .Height) / 2 + Margin
                    .Caption = sTitle
                    .Width = UserControl.TextWidth(.Caption) + 300
                    
                    If Len(sDefaultValue) > 0 Then
                        .Value = 1
                        Combo1.Enabled = True
                    Else
                        .Value = 0
                        Combo1.Enabled = False
                    End If
                    
                    .TabIndex = 0: Seq = 1
                    .Visible = True
                    lblTitle.Visible = False
                    
                    LeftStartMargin = .Left + .Width + 100
                End With
            Else
                With lblTitle
                    .Left = Margin: .Top = (Combo1.Height - .Height) / 2 + Margin
                    .Caption = sTitle
                    .Visible = True
                    Check1(0).Visible = False
                    
                    LeftStartMargin = .Left + .Width + 100
                End With
            End If
            With Combo1
                .Left = LeftStartMargin
                .Top = Margin
                
                .TabIndex = Seq
                .Visible = True
            End With
            With lblUnit
                .Left = Combo1.Left + Combo1.Width + 100: .Top = (Combo1.Height - .Height) / 2 + Margin
                .Caption = sUnit
                .Visible = True
            End With
            UserControl.Width = lblUnit.Left + lblUnit.Width + Margin
            UserControl.Height = Combo1.Height + 2 * Margin
        Case iViewMethod = 3 Or iType = 3 '复选
            iValueNum = -1
            iValueNum = UBound(aItemValues)
            If iType = 3 Or iValueNum = -1 Then
                With Check1(0)
                    .Left = Margin
                    .Top = Margin
                    .Caption = sTitle
                    .Width = UserControl.TextWidth(.Caption) + 300
                    UserControl.Width = .Left + .Width + Margin
                    UserControl.Height = .Height + 2 * Margin
                    .TabIndex = 0
                    .Visible = True
                End With
            Else
                With lblTitle
                    .Left = Margin: .Top = (Check1(0).Height - .Height) / 2 + Margin
                    .Caption = sTitle
                    .Visible = True
                    
                    LeftMargin = .Left + .Width + 70
                End With
                
                iMaxWidth = UserControl.Width
                For i = 0 To iValueNum
                    With Check1(i + 1)
                        .Caption = aItemValues(i)
                        .Width = UserControl.TextWidth(.Caption) + 300
                        .TabIndex = i
                        
                        '处理选项的排列
                        If i = 0 Then
                            .Left = LeftMargin + 30
                            .Top = Margin
                        Else
                            If LeftMargin + 30 + .Width + Margin > iMaxWidth Then
                                .Left = lblTitle.Left + lblTitle.Width + 100
                                .Top = Check1(i).Top + Check1(i).Height + Margin
                            Else
                                .Left = LeftMargin + 30
                                .Top = Check1(i).Top
                            End If
                        End If
                        
                        .Visible = True
                    
                        LeftMargin = .Left + .Width
                        If .Left + .Width + Margin > iMaxWidth Then iMaxWidth = .Left + .Width + Margin
                    End With
                Next
                UserControl.Width = iMaxWidth
                UserControl.Height = Check1(Check1.Count - 1).Top + Check1(Check1.Count - 1).Height + Margin
           End If
        Case iViewMethod = 4 And iType <> 3 '单选
            iValueNum = -1
            iValueNum = UBound(aItemValues)
            If iValueNum = -1 Then
                With Check1(0)
                    .Left = Margin
                    .Top = Margin
                    .Caption = sTitle
                    .Width = UserControl.TextWidth(.Caption) + 300
                    UserControl.Width = .Left + .Width + Margin
                    UserControl.Height = .Height + 2 * Margin
                    .TabIndex = 0
                    .Visible = True
                End With
            Else
                With lblTitle
                    .Left = Margin: .Top = (Option1(0).Height - .Height) / 2 + Margin
                    .Caption = sTitle
                    .Visible = True
                    
                    LeftMargin = .Left + .Width + 70
                End With
                
                iMaxWidth = UserControl.Width
                For i = 0 To iValueNum
                    With Option1(i + 1)
                        .Caption = aItemValues(i)
                        .Width = UserControl.TextWidth(.Caption) + 300
                        .TabIndex = i
                        
                        '处理选项的排列
                        If i = 0 Then
                            .Left = LeftMargin + 30
                            .Top = Margin
                        Else
                            If LeftMargin + 30 + .Width + Margin > iMaxWidth Then
                                .Left = lblTitle.Left + lblTitle.Width + 100
                                .Top = Option1(i).Top + Option1(i).Height + Margin
                            Else
                                .Left = LeftMargin + 30
                                .Top = Option1(i).Top
                            End If
                        End If
                        
                        .Visible = True
                    
                        LeftMargin = .Left + .Width
                        If .Left + .Width + Margin > iMaxWidth Then iMaxWidth = .Left + .Width + Margin
                    End With
                Next
                UserControl.Width = iMaxWidth
                UserControl.Height = Option1(Option1.Count - 1).Top + Option1(Option1.Count - 1).Height + Margin
           End If
    End Select
End Sub

Public Property Get AllowMask() As Boolean
    AllowMask = bMask
End Property

Public Property Let AllowMask(ByVal vNewValue As Boolean)
    bMask = vNewValue
    
    ShowItem
End Property
'焦点始终在编辑栏
'Private Sub SetCtrlFocus()
'    Dim i As Long, iNum As Long
'    On Error Resume Next
'    Select Case True '表示法
'        Case (iViewMethod = 0 And iType <> 3) Or (iViewMethod = 1 And iType = 1) '文本
'            Text1.SetFocus
'            If Check1(0).Value = 0 Then
'                Check1(0).SetFocus
'            End If
'        Case iViewMethod = 1 And (iType = 0 Or iType = 2) '范围
'            Text1.SetFocus
'            If Check1(0).Value = 0 Then
'                Check1(0).SetFocus
'            End If
'        Case iViewMethod = 2 And iType <> 3 '下拉
'            Combo1.SetFocus
'            If Check1(0).Value = 0 Then
'                Check1(0).SetFocus
'            End If
'        Case Else
'            Check1(0).SetFocus: Check1(1).SetFocus
'            Option1(0).SetFocus: Option1(1).SetFocus
'    End Select
'End Sub

Public Property Get TextMethod() As Integer
    TextMethod = iTextMethod
End Property

Public Property Get NullString() As String
    NullString = sNullString
End Property

Public Property Get Unit() As String
    Unit = sUnit
End Property

Public Property Let Unit(ByVal vNewValue As String)
    sUnit = vNewValue
    
    ShowItem
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
