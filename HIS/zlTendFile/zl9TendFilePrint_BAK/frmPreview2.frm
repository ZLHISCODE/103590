VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreview2 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "´òÓ¡Ô¤ÀÀ"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmPreview2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "´òÓ¡"
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bigsmall"
            Object.ToolTipText     =   "ËõÐ¡/·Å´ó"
            ImageKey        =   "bigsmall"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "first"
            Object.ToolTipText     =   "µÚÒ»Ò³"
            ImageKey        =   "first"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previous"
            Object.ToolTipText     =   "ÉÏÒ»Ò³"
            ImageKey        =   "previous"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "next"
            Object.ToolTipText     =   "ÏÂÒ»Ò³"
            ImageKey        =   "next"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "last"
            Object.ToolTipText     =   "×îºóÒ³"
            ImageKey        =   "last"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "°ïÖú"
            ImageKey        =   "help"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "close"
            Object.ToolTipText     =   "¹Ø±Õ"
            ImageKey        =   "close"
         EndProperty
      EndProperty
      Begin VB.TextBox txtÒ³Âë 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   270
         Left            =   5115
         TabIndex        =   7
         Text            =   "txtÒ³Âë"
         Top             =   45
         Width           =   1785
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5055
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":030A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":0464
            Key             =   "tosmall"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":05BE
            Key             =   "tobig"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":0718
            Key             =   "first"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":0B6A
            Key             =   "help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":0CC4
            Key             =   "last"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":1116
            Key             =   "next"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":1568
            Key             =   "bigsmall"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":16C2
            Key             =   "previous"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview2.frx":1B14
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   7005
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "frmPreview2.frx":1C6E
            Text            =   "Á¬½Óµ½LPT1: µÄ Star AR3200"
            TextSave        =   "Á¬½Óµ½LPT1: µÄ Star AR3200"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4948
            MinWidth        =   4939
            Text            =   "ºáÏò   A4£¬210x297ºÁÃ×"
            TextSave        =   "ºáÏò   A4£¬210x297ºÁÃ×"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2990
            Text            =   "Ö½ÕÅÀ´Ô´:"
            TextSave        =   "Ö½ÕÅÀ´Ô´:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "ÏÔÊ¾±ÈÀý£º"
            TextSave        =   "ÏÔÊ¾±ÈÀý£º"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   100
      Left            =   180
      Max             =   1000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2670
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.TextBox txt¶¥°å 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Height          =   440
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   1740
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1815
      LargeChange     =   100
      Left            =   1560
      Max             =   1000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   855
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt¸ô°å 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1695
      Left            =   2115
      TabIndex        =   3
      Top             =   930
      Width           =   255
   End
   Begin VB.PictureBox PctPage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3090
      Index           =   0
      Left            =   2160
      MouseIcon       =   "frmPreview2.frx":1DC8
      MousePointer    =   99  'Custom
      ScaleHeight     =   3060
      ScaleWidth      =   2550
      TabIndex        =   0
      Top             =   1320
      Width           =   2580
      Begin VB.Image imgPage 
         Height          =   3015
         Left            =   120
         MouseIcon       =   "frmPreview2.frx":1F12
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2505
      End
   End
   Begin VB.Shape shpµ×°æ 
      BorderColor     =   &H00000000&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   3330
      Left            =   2880
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmPreview2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnAskPrint As Boolean

Private Sub Form_Activate()
    blnAskPrint = False
    HScroll.Left = 0
    HScroll.Top = ScaleHeight - StatusBar.Height - HScroll.Height
    HScroll.Height = 255
    HScroll.Width = ScaleWidth - HScroll.Left - VScroll.Width
    
    VScroll.Top = 415
    VScroll.Left = ScaleWidth - VScroll.Width
    VScroll.Height = ScaleHeight - StatusBar.Height - VScroll.Top - HScroll.Height
    VScroll.Width = 255
    
    txt¶¥°å.Top = 0
    txt¶¥°å.Left = 0
    txt¶¥°å.Height = 440
    txt¶¥°å.Width = ScaleWidth
    
    
    txt¸ô°å.Top = 0
    txt¸ô°å.Left = VScroll.Left
    txt¸ô°å.Height = ScaleHeight
    txt¸ô°å.Width = 255

    
    txtÒ³Âë.Left = ScaleWidth - txtÒ³Âë.Width - VScroll.Width
    txtÒ³Âë.Text = "¹²" & PctPage.Count - 1 & "Ò³     µÚ1Ò³"
    
    Dim iPaper
    For iPaper = 0 To PctPage.Count - 1
        PctPage(iPaper).Visible = False
    Next
    
    If PctPage(1).Width / PctPage(1).Height > HScroll.Width / VScroll.Height Then
        PctPage(0).Width = HScroll.Width
        PctPage(0).Height = PctPage(1).Height * PctPage(0).Width / PctPage(1).Width
    Else
        PctPage(0).Height = VScroll.Height
        PctPage(0).Width = PctPage(1).Width * PctPage(0).Height / PctPage(1).Height
    End If
    
    If PctPage(1).Width <= PctPage(0).Width Or PctPage(1).Height <= PctPage(0).Height Then
        Toolbar.Buttons("bigsmall").Visible = False
        PctPage(0).Enabled = False
        PctPage(0).MousePointer = 0
        PctPage(1).Enabled = False
        PctPage(1).MousePointer = 0
    End If
    imgPage.Top = 0
    imgPage.Left = 0
    imgPage.Height = PctPage(0).Height
    imgPage.Width = PctPage(0).Width
    
    Tag = 1
    InitScr Tag
    
End Sub

Private Sub InitScr(ByVal iPageNo As Integer)
    
    PctPage(iPageNo).Visible = True
    PctPage(iPageNo).Left = 90
    PctPage(iPageNo).Top = 495
    
    If PctPage(iPageNo).Height > VScroll.Height Then
        VScroll.Visible = True
        txt¸ô°å.Visible = True
    Else
        VScroll.Visible = False
        txt¸ô°å.Visible = False
        PctPage(iPageNo).Top = PctPage(iPageNo).Top + (VScroll.Height - PctPage(iPageNo).Height) / 2
    End If
    
    If PctPage(iPageNo).Width > HScroll.Width Then
        HScroll.Visible = True
    Else
        HScroll.Visible = False
        PctPage(iPageNo).Left = PctPage(iPageNo).Left + (HScroll.Width - PctPage(iPageNo).Width) / 2
    End If
    
    If Not VScroll.Visible And HScroll.Visible Then
        HScroll.Width = ScaleWidth - HScroll.Left
    End If
    If VScroll.Visible And Not HScroll.Visible Then
        VScroll.Height = ScaleHeight - StatusBar.Height - VScroll.Top
    End If
    
    If VScroll.Visible Then
        VScroll.Min = 0
        VScroll.Tag = 0
        VScroll.Value = 0
        VScroll.Max = PctPage(iPageNo).Height - VScroll.Height + 180
        VScroll.SmallChange = PctPage(iPageNo).Height / 200
        VScroll.LargeChange = PctPage(iPageNo).Height / 10
    End If
    
    If HScroll.Visible Then
        HScroll.Min = 0
        HScroll.Tag = 0
        HScroll.Value = 0
        HScroll.Max = PctPage(iPageNo).Width - HScroll.Width + 360
        HScroll.SmallChange = PctPage(iPageNo).Height / 200
        HScroll.LargeChange = PctPage(iPageNo).Height / 10
        HScroll.Tag = HScroll.Value
    End If
    
    shpµ×°æ.Left = PctPage(iPageNo).Left + 45
    shpµ×°æ.Top = PctPage(iPageNo).Top + 45
    shpµ×°æ.Height = PctPage(iPageNo).Height
    shpµ×°æ.Width = PctPage(iPageNo).Width

    StatusBar.Panels(1).Text = "Á¬½Óµ½" & Printer.Port _
        & " µÄ " & Printer.DeviceName
    StatusBar.Panels(2).Text = PaperName()
    StatusBar.Panels(3).Text = "½øÖ½·½Ê½£º" & PaperSource()
    StatusBar.Panels(4).Text = "ÏÔÊ¾±ÈÀý£º" _
        & Int(PctPage(iPageNo).Width / Printer.Width * 100) & "%"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub

Private Sub imgPage_DblClick()
    PctPage(0).Visible = False
    InitScr Tag
End Sub

Private Sub PctPage_DblClick(Index As Integer)
    PctPage(0).Visible = True
    imgPage.Picture = PctPage(Tag).Image
    PctPage(Tag).Visible = False
    InitScr 0
    Refresh
End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim OldPage
    Select Case LCase(Button.Key)
    Case "print"
        blnAskPrint = True
        Me.Hide
    Case "bigsmall"
        If PctPage(0).Visible Then
            PctPage(0).Visible = False
            InitScr Tag
        Else
            PctPage_DblClick 0
        End If
    Case "first", "previous", "next", "last"
        Select Case LCase(Button.Key)
        Case "first"
            If Tag = 1 Then Exit Sub
            OldPage = Tag
            Tag = 1
        Case "previous"
            If Tag = 1 Then Exit Sub
            OldPage = Tag
            Tag = Tag - 1
        Case "next"
            If Tag = PctPage.Count - 1 Then Exit Sub
            OldPage = Tag
            Tag = Tag + 1
        Case "last"
            If Tag = PctPage.Count - 1 Then Exit Sub
            OldPage = Tag
            Tag = PctPage.Count - 1
        End Select
        
        If PctPage(0).Visible Then
            imgPage.Picture = PctPage(Tag).Image
        Else
            PctPage(Tag).Left = PctPage(OldPage).Left
            PctPage(Tag).Top = PctPage(OldPage).Top
            PctPage(Tag).Visible = True
            PctPage(OldPage).Visible = False
        End If
        txtÒ³Âë.Text = "¹²" & PctPage.Count - 1 & "Ò³     µÚ" & Tag & "Ò³"
        Refresh
    Case "help"
        
    Case "close"
        Unload Me
    End Select

End Sub

Private Sub HScroll_Change()
    shpµ×°æ.Move shpµ×°æ.Left - HScroll.Value + HScroll.Tag
    PctPage(Tag).Move PctPage(Tag).Left - HScroll.Value + HScroll.Tag
    HScroll.Tag = HScroll.Value
End Sub

Private Sub HScroll_Scroll()
    shpµ×°æ.Move shpµ×°æ.Left - HScroll.Value + HScroll.Tag
    PctPage(Tag).Move PctPage(Tag).Left - HScroll.Value + HScroll.Tag
    HScroll.Tag = HScroll.Value
End Sub


Private Sub VScroll_Change()
    shpµ×°æ.Move shpµ×°æ.Left, shpµ×°æ.Top - VScroll.Value + VScroll.Tag
    PctPage(Tag).Move PctPage(Tag).Left, PctPage(Tag).Top - VScroll.Value + VScroll.Tag
    VScroll.Tag = VScroll.Value
End Sub

Private Sub VScroll_Scroll()
    shpµ×°æ.Move shpµ×°æ.Left, shpµ×°æ.Top - VScroll.Value + VScroll.Tag
    PctPage(Tag).Move PctPage(Tag).Left, PctPage(Tag).Top - VScroll.Value + VScroll.Tag
    VScroll.Tag = VScroll.Value
End Sub
