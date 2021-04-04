VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl WordInputText 
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   ScaleHeight     =   5295
   ScaleWidth      =   8535
   Begin RichTextLib.RichTextBox txtWord 
      Height          =   4215
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7435
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"WordInputText.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin zl9PACSWork.WordInputModule wimModule 
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7435
      CurDepartId     =   0
   End
   Begin VB.PictureBox picNull 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   6480
      ScaleHeight     =   855
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   120
      Width           =   15
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   4080
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "WordInputText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mblnReadOnly As Boolean
'Private mlngWordWidth As Long



'Property Get ModuleHeight() As Long
'    ModuleHeight = wimModule.WordHeight
'End Property


'Property Let ModuleHeight(value As Long)
'     wimModule.WordHeight = value
'End Property




'Property Get WordWidth() As Long
'    WordWidth = mlngWordWidth
'End Property
'
'
'Property Let WordWidth(value As Long)
'    If mlngWordWidth = value Then Exit Property
'
'    mlngWordWidth = value
'
'    Call InitFace
'End Property



'只读属性
Property Get ReadOnly() As Boolean
    ReadOnly = mblnReadOnly
End Property


Property Let ReadOnly(value As Boolean)
    If mblnReadOnly = value Then Exit Property
    
    mblnReadOnly = value
    
    If mblnReadOnly Then
        txtWord.Locked = True
        
        txtWord.BackColor = UserControl.BackColor
    Else
        txtWord.Locked = False
    
        txtWord.BackColor = vbWhite
    End If
    
End Property






'词库模板名称
Property Get ModuleName() As String
    ModuleName = wimModule.ModuleName
End Property


Property Let ModuleName(value As String)
    wimModule.ModuleName = value
End Property


'当前部门ID
Property Get CurDepartId() As Long
    CurDepartId = wimModule.CurDepartId
End Property

Property Let CurDepartId(value As Long)
    wimModule.CurDepartId = value
End Property



Property Get WordText() As String
    WordText = txtWord.Text
End Property

Property Let WordText(value As String)
    txtWord.Text = value
End Property

Public Sub LoadWordModel()
'载入词库模板
    Call wimModule.LoadWordModel
End Sub




Private Sub InitFace()
'初始化界面布局
    Dim Pane1 As Pane, Pane2 As Pane
    Dim lngWordWidth As Long
    
    With dkpMain
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    lngWordWidth = Round(Width / 3) * 2
    
    Set Pane1 = dkpMain.CreatePane(1, Width - lngWordWidth, 0, DockLeftOf, Nothing)
    Pane1.Title = "词句模板"
    Pane1.Handle = wimModule.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Pane1.MinTrackSize.Width = 50
    Pane1.MinTrackSize.Height = 50
    
    Set Pane2 = dkpMain.CreatePane(2, lngWordWidth, 0, DockRightOf, Pane1)
    Pane2.Title = "词句编辑"
    Pane2.Handle = txtWord.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Pane2.MinTrackSize.Width = 50
    Pane2.MinTrackSize.Height = 50
End Sub



Private Sub UserControl_Initialize()
    mblnReadOnly = False
'    mlngWordWidth = Round(Width / 3) * 2
    
    
    Call InitFace
End Sub



Private Sub wimModule_OnWordDbClickEvent(ByVal strWord As String)
    If Not txtWord.Locked Then txtWord.SelText = strWord
End Sub







Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    wimModule.ModuleName = PropBag.ReadProperty("WordModule", "")
    wimModule.CurDepartId = PropBag.ReadProperty("DepartId", "")
    
    ReadOnly = PropBag.ReadProperty("ReadOnly", False)
    
'    WordWidth = PropBag.ReadProperty("WordWidth", Round(Width / 3) * 2)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next

    Call PropBag.WriteProperty("Module", wimModule.ModuleName, "")
    Call PropBag.WriteProperty("DepartId", wimModule.CurDepartId, "")
    
    Call PropBag.WriteProperty("ReadOnly", mblnReadOnly, False)
    
'    Call PropBag.WriteProperty("WordWidth", mlngWordWidth, Round(Width / 3) * 2)
End Sub


