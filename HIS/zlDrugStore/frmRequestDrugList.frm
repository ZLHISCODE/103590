VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmRequestDrugList 
   Caption         =   "ҩƷ�������"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmRequestDrugList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   3615
      TabIndex        =   11
      Top             =   4200
      Width           =   3615
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1680
         TabIndex        =   17
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   37
         Width           =   720
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   2640
         TabIndex        =   15
         Top             =   30
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1005
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   6255
      _cx             =   11033
      _cy             =   1773
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRequestDrugList.frx":014A
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   180
      MousePointer    =   7  'Size N S
      ScaleHeight     =   360
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   2700
      Width           =   4815
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "��۽�"
         Height          =   180
         Left            =   3690
         TabIndex        =   8
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽�"
         Height          =   180
         Left            =   1890
         TabIndex        =   7
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ���"
         Height          =   180
         Left            =   0
         TabIndex        =   6
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��Χ:1999��8��12����1999��9��12��"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   200
         Width           =   3690
      End
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11775
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinWidth1       =   6000
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "ҩ��"
      Child2          =   "cboStock"
      MinWidth2       =   3000
      MinHeight2      =   300
      Width2          =   3345
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   8685
         TabIndex        =   5
         Text            =   "cboStock"
         Top             =   240
         Width           =   3000
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "PrintView"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Hank"
                     Text            =   "�ֹ���д"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Auto"
                     Text            =   "�Զ�����"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Sale"
                     Text            =   "�Զ�������������"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Key             =   "Merge"
                     Text            =   "�ϲ����쵥"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Receive"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "DisReceive"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmRequestDrugList.frx":01BF
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4620
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRequestDrugList.frx":04D9
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11880
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":0D6D
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":0F8D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":11AD
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":13C9
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":15E9
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1809
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1A25
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1C41
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1E5B
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1FB5
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":21D1
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   600
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":23F1
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2611
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2831
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2A4D
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2C6D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2E8D
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":30A9
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":32C5
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":34DF
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":3639
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":3859
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   5655
      _cx             =   9975
      _cy             =   1720
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
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRequestDrugList.frx":3A79
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "���ݴ�ӡ(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileBillPreview 
         Caption         =   "����Ԥ��(&L)"
      End
      Begin VB.Menu mnuFileCodePrint 
         Caption         =   "�����ӡ(&C)"
         Begin VB.Menu mnuFileAllCodePrint 
            Caption         =   "������ҩƷ�����ӡ(&A)"
         End
         Begin VB.Menu mnuFileSelCodePrint 
            Caption         =   "ѡ����ҩƷ�����ӡ(&S)"
         End
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Begin VB.Menu mnuEditAddHank 
            Caption         =   "�ֹ���д(&H)"
         End
         Begin VB.Menu mnuEditAddAuto 
            Caption         =   "�Զ�����(&A)"
         End
         Begin VB.Menu mnuEditAddAutoBySale 
            Caption         =   "�Զ�������������(&S)"
         End
         Begin VB.Menu mnuEditAddMerge 
            Caption         =   "�ϲ����쵥(&M)"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditReceive 
         Caption         =   "���(&R)"
      End
      Begin VB.Menu mnuEditDisReceive 
         Caption         =   "����(&D)"
      End
      Begin VB.Menu mnuEditWriteOff 
         Caption         =   "��������(&W)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "�鿴����(&W)"
      End
      Begin VB.Menu mnuEditCodePrintLine 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditAllCodePrint 
         Caption         =   "������ҩƷ�����ӡ(&A)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmRequestDrugList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String
Private mblnBootUp As Boolean
Private mlastRow As Long                '�ϴε������
Private mintPreCol As Integer           'ǰһ�ε���ͷ��������
Private mintsort As Integer             'ǰһ�ε���ͷ������
Private mintPreDetailCol As Integer     'ǰһ�ε������������
Private mintDetailsort As Integer       'ǰһ�ε����������
Private mlngMode As Long
Private mstrPrivs As String             '��ǰ�û����еĵ�ǰģ��Ĺ���
Private mint��ѯ���� As Integer
Private mblnViewCost As Boolean     '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���
Private Const MStrCaption As String = "ҩƷ�������"

Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    lng�Ƴ��ⷿ As Long
    str������ As String
    str����� As String
End Type

Private SQLCondition As Type_SQLCondition

Private mlng�ⷿid As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private mint�������� As Integer                          '0-����Ҫ����;1-��Ҫ����

'�Ӳ�������ȡҩƷ�۸����������С��λ������ʾ���ȣ�
Private mintShowCostDigit As Integer            '�ɱ���С��λ��
Private mintShowPriceDigit As Integer           '�ۼ�С��λ��
Private mintShowNumberDigit As Integer          '����С��λ��
Private mintShowMoneyDigit As Integer           '���С��λ��

Private mstrNumberFormat As String
Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrMoneyFormat As String

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Public Function CheckBill(ByVal strNo As String) As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHandle
    CheckBill = ""
    
    gstrSQL = " Select �������,��ҩ����,��ҩ�� From ҩƷ�շ���¼ " & _
            " Where ����=6 And NO=[1] And ��¼״̬=1 And RowNum=1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, "��鵥��", strNo)
    
    With rs
        '���ؿգ���ʾ�Ѿ�ɾ��
        If .EOF Then
            CheckBill = "�õ����Ѿ�����������Աɾ����"
        ElseIf Not IsNull(!�������) Then
            CheckBill = "�õ����Ѿ�����������Ա��ˣ�"
        ElseIf Not IsNull(!��ҩ����) Then
            CheckBill = "�õ����Ѿ�����������Ա���ͣ�"
        ElseIf Not IsNull(!��ҩ��) Then
            CheckBill = "�õ����Ѿ�����������Ա��ҩ��"
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckNoIsExist(ByVal StrBillNo As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '��鵥���Ƿ����
    On Error GoTo errHandle
    gstrSQL = " Select id From ҩƷ�շ���¼ " & _
              " Where ����=6 And NO=[1] And ���ϵ�� = -1 and rownum = 1"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "��鵥���Ƿ����", StrBillNo)
    CheckNoIsExist = Not (rsCheck.RecordCount = 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub cboStock_Click()
    If mlng�ⷿid <> cboStock.ItemData(cboStock.ListIndex) Then
        mlng�ⷿid = cboStock.ItemData(cboStock.ListIndex)
        Call GetDrugDigit(mlng�ⷿid, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '��֯��ʽ����
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
        If mblnBootUp Then mnuViewRefresh_Click
    End If
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    str�������� = "H,I,J,K,L,M,N"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboStock, Trim(cboStock.Text), str��������, True, "0,1,2,3") = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub


Private Sub cboStock_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    If cboStock.ListCount > 0 Then
        If cboStock.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub


Private Sub cbrTool_Resize()
    If mblnBootUp = False Then Exit Sub
    Form_Resize
End Sub

Public Sub ShowList(ByVal FrmMain As Variant)
    Dim strFind As String
    Dim dateCurDate As Date
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs

    mblnBootUp = False
    
    dateCurDate = Sys.Currentdate()
    mint��ѯ���� = Val(zlDataBase.GetPara("��ѯ����", glngSys, 1343)) - 1
    
    'ȡ�ƿ�ĳ����������
    mint�������� = Val(zlDataBase.GetPara("��������", glngSys, 1304))
    
    strStart = Format(DateAdd("d", -1 * mint��ѯ����, dateCurDate), "yyyy-MM-dd")
    strEnd = Format(dateCurDate, "yyyy-MM-dd")
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
        
    If Not CheckDepend Then Exit Sub            '���������Բ���
    
    mlng�ⷿid = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng�ⷿid, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
    
    strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4] "
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = strFind
    
    GetList (mstrFind)  '�г�����ͷ
    RestoreWinState Me, App.ProductName, MStrCaption
    Call zlDataBase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    If mblnViewCost = False Then
        vsfDetail.ColWidth(11) = 0
        vsfDetail.ColWidth(12) = 0
    End If
    
    mblnBootUp = True
        
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        'ZLBH�ںϵ���
        OS.ShowChildWindow Me.hWnd, FrmMain
    End If
    Me.ZOrder 0
End Sub

'�������������
Private Function CheckDepend() As Boolean
    
    Dim rsDepend As New Recordset
    Dim strStock As String
    
    On Error GoTo errHandle
    CheckDepend = False
    strStock = "HIJKLMN"
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
             & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where c.�������� = b.���� " _
              & "AND Instr([1],b.����,1) > 0 " _
             & " AND a.id = c.����id " _
              & "AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"

    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�������", strStock)
    
    If rsDepend.EOF Then
        MsgBox "����������Ϣ��ȫ,��鿴���Ź���", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
            
    rsDepend.Close

    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
         & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
        & "Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� is Null) And c.�������� = b.���� " _
          & "AND Instr([1],b.����,1) > 0 " _
         & " AND a.id = c.����id " _
         & " and a.id in (select ����id from ������Ա where ��Աid= [2]) " _
          & " AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�������", strStock, glngUserId)
    
    If rsDepend.EOF Then
        MsgBox "�㲻��ҩ�⡢ҩ�����Ƽ��ҵĹ�����Ա�����ܽ��룡", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!����
            .ItemData(.NewIndex) = rsDepend!Id
            If rsDepend!Id = glngDeptId Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        
        If .ListIndex = -1 Then
            .ListIndex = 0
        End If
    End With
    
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList(ByVal strFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    
    '����ͳ�ƺϼƽ��
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim n As Long
    Dim strFormat As String
    
    On Error GoTo errHandle
    strFormat = "0.00##"
    
    mlastRow = 0
    
    vsfList.Redraw = flexRDNone
    strUserPart = " And A.�ⷿID+0=[11] "
    gstrSQL = "SELECT A.NO, C.���� AS ��ҩ�ⷿ,LTRIM(TO_CHAR (SUM (A.�ɱ����)," & mstrCostFormat & ")) AS �ɱ����, " & _
        " LTRIM(TO_CHAR ( (SUM (A.���۽��)), " & mstrMoneyFormat & ")) AS �ۼ۽��,LTRIM(TO_CHAR (SUM (A.���۽�� - A.�ɱ����), " & mstrMoneyFormat & ")) AS ��۽��, A.������, " & _
        " TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������,A.�޸���,TO_CHAR (MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����, A.�����, " & _
        " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.��ҩ�� ������,A.ժҪ " & _
        " FROM ҩƷ�շ���¼ A, ���ű� B,���ű� C " & _
        " WHERE A.�ⷿID = B.ID AND A.�Է�����ID=C.ID AND A.���� = 6 AND  A.���ϵ��=1 " & _
        " And (A.��ҩ�� Is NULL Or A.��ҩ���� Is Not NULL)" & _
        strUserPart & strFind & _
        " GROUP BY A.NO,C.����,A.������,A.�޸���,A.�����,A.��¼״̬,A.��ҩ��,A.ժҪ " & _
        " ORDER BY NO DESC, �������� ASC "
    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
        SQLCondition.strNO��ʼ, _
        SQLCondition.strNO����, _
        SQLCondition.date����ʱ�俪ʼ, _
        SQLCondition.date����ʱ�����, _
        SQLCondition.date���ʱ�俪ʼ, _
        SQLCondition.date���ʱ�����, _
        SQLCondition.lngҩƷ, _
        SQLCondition.lng�Ƴ��ⷿ, _
        SQLCondition.str������, _
        SQLCondition.str�����, _
        cboStock.ItemData(cboStock.ListIndex))
        
    Set vsfList.DataSource = rsList
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = flexRDDirect
            
            .TopRow = 1
            .rows = .rows - 99
        End If
        .Row = 1
        .Col = 0
        For n = 0 To .Cols - 1
            .ColKey(n) = .TextMatrix(0, n)
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
        '.ColSel = .Cols - 1    ' bug: 40410
    End With
    SetListColWidth
    
    'ͳ�ƺϼƽ��
    If (Not rsList.EOF) And (Not rsList.BOF) Then
        rsList.MoveFirst
        Do While Not rsList.EOF
            dbl1 = dbl1 + IIf(IsNull(rsList!�ɱ����), 0, rsList!�ɱ����)
            dbl2 = dbl2 + IIf(IsNull(rsList!�ۼ۽��), 0, rsList!�ۼ۽��)
            dbl3 = dbl3 + IIf(IsNull(rsList!��۽��), 0, rsList!��۽��)
            rsList.MoveNext
        Loop
        rsList.MoveFirst
        
        lbl1.Caption = "�ɱ����ϼƣ�" & zlStr.FormatEx(dbl1, mintShowMoneyDigit, , True)
        lbl2.Caption = "�ۼ۽��ϼƣ�" & zlStr.FormatEx(dbl2, mintShowMoneyDigit, , True)
        lbl3.Caption = "��۽��ϼƣ�" & zlStr.FormatEx(dbl3, mintShowMoneyDigit, , True)

    End If
    vsfList_EnterCell    '�г�������
    
    SetStrikeColor
    staThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    rsList.Close
    vsfList.Redraw = flexRDDirect
        
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
        
    With vsfList
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
            intStatus = .TextMatrix(intRow, .Cols - 3)
            If intStatus Mod 3 = 0 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001
            End If
            If intStatus Mod 3 = 2 Then
                '�ƿ��������������Ϊǳ��ɫ���ѳ�������Ϊ��ɫ
                If Trim(.TextMatrix(intRow, GetCol(vsfList, "�������"))) <> "" Then
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF
                Else
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF00FF       ' &HC0C0FF
                End If
            End If
        Next
    End With
                
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter

        For intCol = 1 To .Cols - 1
            If intCol = 1 Then
               .ColWidth(intCol) = 2000
            ElseIf intCol = .Cols - 3 Then
                .ColWidth(intCol) = 0
            Else
                .ColWidth(intCol) = 1000
            End If
        Next
        If mblnViewCost = False Then
            .ColHidden(.ColIndex("�ɱ����")) = True
            .ColHidden(.ColIndex("��۽��")) = True
        End If
    End With
End Sub


Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    On Error GoTo errHandle
    
    str�ⷿ���� = ""
    gstrSQL = "Select a.�������� From ��������˵�� A Where a.����id =[1]"
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsDetail.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
        rsDetail.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
        
    With vsfDetail
        .ColAlignment(.ColIndex("��д����")) = flexAlignRightCenter     '��д����
        .ColAlignment(.ColIndex("ʵ������")) = flexAlignRightCenter     'ʵ������
        .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
        .ColAlignment(.ColIndex("�ɱ���")) = flexAlignRightCenter     '�ɱ���
        .ColAlignment(.ColIndex("�ɱ����")) = flexAlignRightCenter     '�ɱ����
        .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter    '�ۼ�
        .ColAlignment(.ColIndex("�ۼ۽��")) = flexAlignRightCenter    '�ۼ۽��
        .ColAlignment(.ColIndex("���")) = flexAlignRightCenter    '���
                
        .ColWidth(0) = 0
        .ColWidth(.ColIndex("ҩƷ��Ϣ")) = 2500
        For intCol = 2 To .Cols - 1
            .ColWidth(intCol) = 1000
        Next
        If mblnViewCost = False Then
            .ColHidden(.ColIndex("�ɱ���")) = True
            .ColHidden(.ColIndex("�ɱ����")) = True
            .ColHidden(.ColIndex("���")) = True
        End If
        
        If bln��ҩ�ⷿ Then
            .ColHidden(.ColIndex("ԭ����")) = False
        Else
            .ColHidden(.ColIndex("ԭ����")) = True
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetVisable()
    '����������
    If Not IsHavePrivs(mstrPrivs, "����") Or (Not IsHavePrivs(mstrPrivs, "�ֹ�����") And Not IsHavePrivs(mstrPrivs, "�Զ�����") And Not IsHavePrivs(mstrPrivs, "�������Զ�����")) Then
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDel.Visible = False
        
        tlbTool.Buttons("Add").Visible = False
        tlbTool.Buttons("Modify").Visible = False
        tlbTool.Buttons("Delete").Visible = False
        tlbTool.Buttons("Edit1").Visible = False
        mnuEditLine1.Visible = False
    Else
        If Not IsHavePrivs(mstrPrivs, "�ֹ�����") Then
            mnuEditAddHank.Visible = False
            tlbTool.Buttons("Add").ButtonMenus("Hank").Visible = False
        End If
        If Not IsHavePrivs(mstrPrivs, "�Զ�����") Then
            mnuEditAddAuto.Visible = False
            tlbTool.Buttons("Add").ButtonMenus("Auto").Visible = False
        End If
        If Not IsHavePrivs(mstrPrivs, "�������Զ�����") Then
            mnuEditAddAutoBySale.Visible = False
            tlbTool.Buttons("Add").ButtonMenus("Sale").Visible = False
        End If
    End If
    If Not IsHavePrivs(mstrPrivs, "���") Then
        mnuEditReceive.Visible = False
        tlbTool.Buttons("Receive").Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "����") Then
        mnuEditDisReceive.Visible = False
        mnuEditWriteOff.Visible = False
        If mnuEditReceive.Visible = False Then mnuEditLine2.Visible = False
        tlbTool.Buttons("DisReceive").Visible = False
        tlbTool.Buttons("EditSeparate").Visible = mnuEditLine2.Visible
        mnuEditWriteOff.Visible = False
'        tlbTool.Buttons("DisReceive").ButtonMenus(1).Visible = False
'        tlbTool.Buttons("DisReceive").ButtonMenus(2).Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
    
End Sub

Private Sub Form_Activate()
    If mint�������� = 1 Then
        mnuEditDisReceive.Caption = "�������(&R)"
        tlbTool.Buttons("DisReceive").Caption = "�������"
    Else
        mnuEditDisReceive.Caption = "����(&D)"
        tlbTool.Buttons("DisReceive").Caption = "����"
    End If
End Sub

Private Sub Form_Load()
    '�ָ�����
    mblnViewCost = IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        lblRange.Visible = False
    ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        lblRange.Visible = True
        lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
        lblRange.Visible = True
        lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
    ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        lblRange.Visible = True
        lblRange = "��ѯ��Χ:������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    End If
    
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    lbl2.Left = lbl1.Left + lbl1.Width + 3000
    lbl3.Left = lbl2.Left + lbl2.Width + 3000
    If mblnViewCost = False Then
        lbl1.Visible = False
        lbl3.Visible = False
        lbl2.Left = lbl1.Left
    End If
    
    staThis.Panels(2).Picture = picColor
End Sub

Private Sub Form_Resize()
    '����λ������
    
    On Error Resume Next
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picSeparate_s
        .Height = 360
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With vsfDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        .Width = cbrTool.Width
    End With
        
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - .Width - 300
    End With
    
    If mlngMode <> 1343 Then
        picColor3.Visible = False
        lblColor3.Visible = False
        picColor.Width = lblColor2.Left + lblColor2.Width + 20
    Else
        lblColor3.Caption = "δ��˳���"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, MStrCaption
End Sub

Private Sub mnuEditAddAuto_Click()
    '�Զ�����
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '��鱾���Ƿ��Ѿ���˽�棬���δ��˽�����ܽ�������ҵ�����
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    strNo = ""
    frmRequestDrugCard.ShowCard Me, strNo, 5, , BlnSuccess, cboStock.ItemData(cboStock.ListIndex), 0, 1
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditAddAutoBySale_Click()
    '�Զ����������Զ�����
    Dim strNo As String
    Dim BlnSuccess As Boolean
    Dim rsTmp As ADODB.Recordset
    
    '��鱾���Ƿ��Ѿ���˽�棬���δ��˽�����ܽ�������ҵ�����
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    On Error GoTo errHandle
    '��������췽ʽ����Ҫ����Ƿ������
    gstrSQL = "Select 1 From ҩƷ�շ���¼ " & _
        " Where ���� = 6 And �ⷿid = [1] And ���� = 7 And ������� Is Null and �������� between sysdate-60 and sysdate and rownum=1"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "mnuEditAddAutoBySale_Click", Val(cboStock.ItemData(cboStock.ListIndex)))
    
    If Not rsTmp.EOF Then
        MsgBox "������δ��˵��Զ����쵥�ݣ����ܲ����µ��ݡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strNo = ""
    frmRequestDrugCard.ShowCard Me, strNo, 5, , BlnSuccess, cboStock.ItemData(cboStock.ListIndex), 0, 7
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub mnuEditAddHank_Click()
    '�ֹ���д
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '��鱾���Ƿ��Ѿ���˽�棬���δ��˽�����ܽ�������ҵ�����
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    strNo = ""
    '����
    frmRequestDrugCard.ShowCard Me, strNo, 1, , BlnSuccess, cboStock.ItemData(cboStock.ListIndex)
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditAddMerge_Click()
    '�ϲ����쵥
    
'    frmRequestMerge.Show vbModal, Me
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim StrBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
    Dim strCheckString As String
    
    With vsfList
        On Error GoTo errHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, 0)
        
        If Not CheckNoIsExist(StrBillNo) Then
            MsgBox "û���ҵ��õ��ݣ������ѱ�ɾ����", vbInformation, gstrSysName
            
            'ˢ��
            GetList mstrFind
            Exit Sub
        End If
        
        'δ��˵���
        If .TextMatrix(intRow, .Cols - 4) = "" And Val(.TextMatrix(.Row, .Cols - 3)) = 1 Then
            If Not Is����(StrBillNo) Then
                MsgBox "��û��Ȩ��ɾ���ƿⵥ��", vbInformation, gstrSysName
                Exit Sub
            End If
        
            strCheckString = CheckBill(Trim(StrBillNo))
            If strCheckString <> "" Then
                MsgBox strCheckString, vbInformation, gstrSysName
                GetList mstrFind
                Exit Sub
            End If
            
            strTitle = "ҩƷ���쵥"
        ElseIf Val(.TextMatrix(.Row, .Cols - 3)) Mod 3 = 2 And mint�������� = 1 Then
            '����˵��ݣ������������������
            strTitle = "�������뵥"
        End If
        
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & StrBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_ҩƷ�ƿ�_Delete('" & StrBillNo & "'," & Val(.TextMatrix(.Row, .Cols - 3)) & " )"
            If gstrSQL = "" Then Exit Sub
            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption & "-ɾ�����쵥")
            intRecord = intRecord - 1
            mlastRow = 0
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                With vsfDetail
                    .rows = 1
                    .rows = 2
                    .FixedRows = 1
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                SetEnable
                
            End If
                
            '.RowHeight(intRow) = 0
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
            vsfList_EnterCell
        End If
    End With
    staThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    If intRecord = 0 Then
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
    End If
    mnuViewRefresh_Click
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume 'Resume����������õ���
    Call SaveErrLog
    
End Sub

Private Sub mnuEditDisplay_Click()
    '�鿴����
    
    Dim strNo As String
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmRequestDrugCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .Cols - 3), , cboStock.ItemData(cboStock.ListIndex)
    End With
End Sub


'Modified By ���� 2003-12-10 ����������
Private Sub mnuEditDisReceive_Click()
    Dim strNo As String, BlnSuccess As Boolean
    Dim int����ʽ As Integer
    
    If mnuEditDisReceive.Caption = "�������(&R)" Then
        int����ʽ = 1
    Else
        int����ʽ = 0
    End If
                
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmRequestDrugCard.ShowCard Me, strNo, 7, .TextMatrix(.Row, .Cols - 3), BlnSuccess, cboStock.ItemData(cboStock.ListIndex), int����ʽ
        If Not BlnSuccess Then Exit Sub
    End With
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditModify_Click()
    '�޸�
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        If Not CheckNoIsExist(strNo) Then
            MsgBox "û���ҵ��õ��ݣ������ѱ�ɾ����", vbInformation, gstrSysName
            
            'ˢ��
            GetList mstrFind
            Exit Sub
        End If
        
        If Not Is����(strNo) Then
            MsgBox "��û��Ȩ���޸��ƿⵥ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        frmRequestDrugCard.ShowCard Me, strNo, 2, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 3), BlnSuccess, cboStock.ItemData(cboStock.ListIndex)
        If BlnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

'Modified By ���� 2003-12-10 ����������
Private Sub mnuEditReceive_Click()
    Dim strNo As String, BlnSuccess As Boolean
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmRequestDrugCard.ShowCard Me, strNo, 6, .TextMatrix(.Row, .Cols - 3), BlnSuccess, cboStock.ItemData(cboStock.ListIndex)
    End With
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditWriteOff_Click()
    Dim strStock As String
    Dim i As Integer
    
    With Me.cboStock
        For i = 0 To .ListCount - 1
            strStock = strStock & .List(i) & "," & .ItemData(i) & "|"
        Next
       
    End With
    
    Call frm��������.ShowMe(1341, Me, strStock, Me.cboStock.ListIndex)
End Sub

Private Sub mnuFileAllCodePrint_Click()
    If Trim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Or vsfList.rows <= 1 Then Exit Sub
    CodePrint vsfList.TextMatrix(vsfList.Row, 0)
End Sub

Private Sub mnuEditAllCodePrint_Click()
    If Trim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Or vsfList.rows <= 1 Then Exit Sub
    CodePrint vsfList.TextMatrix(vsfList.Row, 0)
End Sub

Private Sub mnuFileSelCodePrint_Click()
    If Trim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Or vsfList.rows <= 1 Then Exit Sub
    CodePrint Val(vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("ҩƷID")))
End Sub

Private Sub CodePrint(ByVal varPar As Variant)
'���ܣ���ӡҪƷ����
'���Σ�varPar��long�����ӡ��ӦҩƷ���룻��String������ݵ��ݺŴ�ӡ�����е�ҩƷ����
    Dim rstemp As New ADODB.Recordset

    On Error GoTo errHandle
    
    If Not IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
        MsgBox "�Բ�����û�и�Ȩ�ޣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TypeName(varPar) = "String" Then '��ӡ���ŵ�������
        gstrSQL = "select distinct ҩƷID from ҩƷ�շ���¼ where ���� = 6 and  NO = [1] order by ҩƷID"
        
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�����ӡ", varPar)
        
        Do While Not rstemp.EOF
            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1343_1", Me, "ҩƷ=" & rstemp!ҩƷID, 2
            rstemp.MoveNext
        Loop
        
    Else '��ӡ��ӦҩƷ����
        If varPar = 0 Then Exit Sub
        ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1343_1", Me, "ҩƷ=" & varPar, 2
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub mnuFileBillPreview_Click()
    Dim int��λϵ�� As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                int��λϵ�� = 4
            Case mconint���ﵥλ
                int��λϵ�� = 2
            Case mconintסԺ��λ
                int��λϵ�� = 1
            Case mconintҩ�ⵥλ
                int��λϵ�� = 3
        End Select
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1304", "zl8_bill_1304"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 3), "��λϵ��=" & int��λϵ��, 1
    End With
End Sub

Private Sub MnuFileBillprint_Click()
    Dim int��λϵ�� As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                int��λϵ�� = 4
            Case mconint���ﵥλ
                int��λϵ�� = 2
            Case mconintסԺ��λ
                int��λϵ�� = 1
            Case mconintҩ�ⵥλ
                int��λϵ�� = 3
        End Select
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1304", "zl8_bill_1304"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 3), "��λϵ��=" & int��λϵ��, 2
    End With
End Sub



Private Sub mnuFileExcel_Click()
    '�����Excel
    If Me.ActiveControl Is vsfList Then
        vsfList.Redraw = flexRDNone
        subPrint 3
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    ElseIf Me.ActiveControl Is vsfDetail Then
        vsfDetail.Redraw = flexRDNone
        subExcel 3
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
End Sub

Private Sub mnufileexit_Click()
    '�˳�
    Unload Me
    
End Sub

Private Sub mnuFileParameter_Click()
    Dim dateCurDate As Date
    
    '��������
    frm��������.���ò��� Me, mstrPrivs, 1343, MStrCaption
    
    Call GetDrugDigit(mlng�ⷿid, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '������֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    dateCurDate = Sys.Currentdate
    mint��ѯ���� = Val(zlDataBase.GetPara("��ѯ����", glngSys, 1343)) - 1
'    mint��ѯ���� = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & MStrCaption, "��ѯ����", "7")
    strStart = Format(DateAdd("d", -1 * mint��ѯ����, dateCurDate), "yyyy-MM-dd")
    strEnd = Format(dateCurDate, "yyyy-MM-dd")
    mstrFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4] "
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        lblRange.Visible = False
    ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        lblRange.Visible = True
        lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
        lblRange.Visible = True
        lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
    ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        lblRange.Visible = True
        lblRange = "��ѯ��Χ:������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    End If
    
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub mnuFilePrintSet_Click()
    '��ӡ����
    zlPrintSet
End Sub


Private Sub mnuHelpAbout_Click()
    '����
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '��������
'    ReportMan gcnOracle, Me
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(index As Integer)
    'Ĭ�ϲ�����ҩƷ=ҩƷid��ҩ��=ҩ��id���ⷿ=�ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬NO=���쵥NO
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim strNo As String
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        strNo = vsfList.TextMatrix(vsfList.Row, 0)
    End If
    
    str��ʼʱ�� = IIf(Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd"))
    str����ʱ�� = IIf(Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd"))
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
        "ҩƷ=" & IIf(SQLCondition.lngҩƷ = 0, "", SQLCondition.lngҩƷ), _
        "ҩ��=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
        "�ⷿ=" & IIf(SQLCondition.lng�Ƴ��ⷿ = 0, "", SQLCondition.lng�Ƴ��ⷿ), _
        "��ʼʱ��=" & str��ʼʱ��, _
        "����ʱ��=" & str����ʱ��, _
        "NO=" & strNo)
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '����
    Dim strFind As String
    
    strFind = FrmListSearch.GetSearch(Me, 1343, cboStock.ItemData(cboStock.ListIndex), strStart, strEnd, strVerifyStart, strVerifyEnd, _
                    SQLCondition.strNO��ʼ, _
                    SQLCondition.strNO����, _
                    SQLCondition.date����ʱ�俪ʼ, _
                    SQLCondition.date����ʱ�����, _
                    SQLCondition.date���ʱ�俪ʼ, _
                    SQLCondition.date���ʱ�����, _
                    SQLCondition.lngҩƷ, _
                    SQLCondition.lng�Ƴ��ⷿ, _
                    SQLCondition.str������, _
                    SQLCondition.str�����)
    
    If strFind <> "" Then
        mstrFind = strFind
        GetList mstrFind
        If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
        ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        End If
             
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '����������
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            'ȡ�����е��ı���ǩ��ʾ
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '�����е��ı���ǩ��ʾ��˵����Tag�зŵ��ı���ǩ
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub

Private Sub vsfDetail_GotFocus()
    Call SetGridFocus(vsfDetail, True)
End Sub

Private Sub vsfDetail_LostFocus()
    Call SetGridFocus(vsfDetail, False)
End Sub

Private Sub vsfDetail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuFileCodePrint.Visible = False Then Exit Sub
    
    PopupMenu mnuFileCodePrint, 2
End Sub

Private Sub vsfList_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If vsfList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub vsfList_EnterCell()
    Dim rsDetail As New Recordset
    Dim strUnitQuantity As String               '��λ��������ʽ����
    Dim IntBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim strUnit As String                       '��λ����:�����ﵥλ��סԺ��λ��
    Dim str��װϵ�� As String
    Dim strOrder As String
    Dim strCompare As String
    Dim strSqlЧ�� As String
    Dim strSqlҩ�� As String
    Dim intCol As Integer
    Dim strSqlOrder As String
    Dim n As Integer
    
    If mlastRow = vsfList.Row Then Exit Sub
    mlastRow = vsfList.Row
    On Error GoTo errHandle
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, mlastRow, 1)
        .Redraw = flexRDDirect
    End With
    
    SetEnable
    
    strOrder = zlDataBase.GetPara("����", glngSys, 1343)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "���"
    
    If strCompare = "0" Then
        '���������
        strSqlOrder = "���"
    ElseIf strCompare = "1" Then
        '����������
        strSqlOrder = "ҩƷ��Ϣ"
    ElseIf strCompare = "2" Then
        '����������
        strSqlOrder = "Substr(ҩƷ��Ϣ, Instr(ҩƷ��Ϣ, ']') + 1)"
    ElseIf strCompare = "3" Then
        ''���ⷿ��λ����
        strSqlOrder = "�ⷿ��λ"
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",ҩƷ��Ϣ,���"
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
        
        vsfDetail.Redraw = flexRDNone
        
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                strUnit = "D.���㵥λ"
                str��װϵ�� = "1"
            Case mconint���ﵥλ
                strUnit = "B.���ﵥλ"
                str��װϵ�� = "B.�����װ"
            Case mconintסԺ��λ
                strUnit = "B.סԺ��λ"
                str��װϵ�� = "B.סԺ��װ"
            Case mconintҩ�ⵥλ
                strUnit = "B.ҩ�ⵥλ"
                str��װϵ�� = "B.ҩ���װ"
        End Select
        
        strSqlЧ�� = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "TO_CHAR(A.Ч��-1,'YYYY-MM-DD') AS ��Ч����", "TO_CHAR(A.Ч��,'YYYY-MM-DD') AS ʧЧ��")
        
        If gintҩƷ������ʾ = 0 Then
            strSqlҩ�� = ",('['||D.����||']'||D.����) AS ҩƷ��Ϣ"
        ElseIf gintҩƷ������ʾ = 1 Then
            strSqlҩ�� = ",('['||D.����||']'||NVL(E.����,D.����)) AS ҩƷ��Ϣ"
        Else
            strSqlҩ�� = ",('['||D.����||']'||D.����) AS ҩƷ��Ϣ,E.���� As ��Ʒ��"
        End If
        
        gstrSQL = " SELECT * FROM (SELECT DISTINCT ���" & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��," & _
            " D.���,A.���� as ������,A.ԭ����, A.����, " & strSqlЧ�� & " ,A.��׼�ĺ�,LTRIM(TO_CHAR(A.��д���� /" & str��װϵ�� & "," & mstrNumberFormat & " )) AS ��д����," & _
            " LTRIM(TO_CHAR(A.ʵ������ /" & str��װϵ�� & "," & mstrNumberFormat & ")) AS ʵ������," & strUnit & " AS ��λ," & _
            " LTRIM(TO_CHAR (A.�ɱ���*" & str��װϵ�� & ", " & mstrCostFormat & ")) AS �ɱ���," & _
            " LTRIM(TO_CHAR (A.�ɱ����, " & mstrMoneyFormat & ")) AS �ɱ����," & _
            " LTRIM(TO_CHAR (A.���ۼ�*" & str��װϵ�� & ", " & mstrPriceFormat & ")) AS �ۼ�," & _
            " LTRIM(TO_CHAR (A.���۽��, " & mstrMoneyFormat & ")) AS �ۼ۽��," & _
            " LTRIM(TO_CHAR (A.���, " & mstrMoneyFormat & ")) AS ��� ,C.�ⷿ��λ,D.ID ҩƷID " & _
            " FROM ҩƷ�շ���¼ A, ҩƷ��� B, �շ���Ŀ���� E, �շ���ĿĿ¼ D, ҩƷ�����޶� C " & _
            " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=D.ID " & _
            " AND B.ҩƷID = E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
            " AND A.��¼״̬ = [2] " & _
            " AND A.���� = 6 AND ���ϵ��=1 " & _
            " AND A.NO = [1] AND A.ҩƷID=C.ҩƷID(+) AND A.�ⷿID=C.�ⷿID(+))" & _
            " ORDER BY " & strSqlOrder
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, vsfList.TextMatrix(vsfList.Row, 0), Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 3)))
        
        Set vsfDetail.DataSource = rsDetail
        rsDetail.Close

        With vsfDetail
            .Row = 1
            .Col = 0
            '.ColSel = .Cols - 1    ' bug: 40410
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
            .ColHidden(.ColIndex("ҩƷID")) = True 'ҩƷID�в���ʾ
        End With
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Redraw = flexRDNone
            .Cols = IIf(gintҩƷ������ʾ = 2, 20, 19)
            .rows = 2
            .Clear
            
            intCol = 0
            
            .TextMatrix(0, intCol) = "���": intCol = intCol + 1
            .TextMatrix(0, intCol) = "ҩƷ��Ϣ": intCol = intCol + 1
            
            If gintҩƷ������ʾ = 2 Then
                .TextMatrix(0, intCol) = "��Ʒ��": intCol = intCol + 1
            End If
                    
            .TextMatrix(0, intCol) = "ҩƷ��Դ": intCol = intCol + 1
            .TextMatrix(0, intCol) = "����ҩ��": intCol = intCol + 1
            .TextMatrix(0, intCol) = "���": intCol = intCol + 1
            .TextMatrix(0, intCol) = "������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "ԭ����": intCol = intCol + 1
            .TextMatrix(0, intCol) = "����": intCol = intCol + 1
            .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
            .TextMatrix(0, intCol) = "��׼�ĺ�": intCol = intCol + 1
            .TextMatrix(0, intCol) = "��д����": intCol = intCol + 1
            .TextMatrix(0, intCol) = "ʵ������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ɱ���": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ɱ����": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ۼ۽��": intCol = intCol + 1
            .TextMatrix(0, intCol) = "���": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ⷿ��λ": intCol = intCol + 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
            .Redraw = flexRDDirect
        End With
    End If
    SetDetailColWidth
    CheckNumber
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_GotFocus()
    Call SetGridFocus(vsfList, True)
End Sub

Private Sub vsfList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub vsfList_LostFocus()
    Call SetGridFocus(vsfList, False)
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    mnuEditAllCodePrint.Visible = True
    mnuEditCodePrintLine.Visible = True
    PopupMenu mnuEdit, 2
    mnuEditAllCodePrint.Visible = False
    mnuEditCodePrintLine.Visible = False
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Height = picSeparate_s.Top - .Top
    End With
    
    With vsfDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
    End With
    
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAddHank_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Receive"
            mnuEditReceive_Click
        Case "DisReceive"
            mnuEditDisReceive_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
    End Select
End Sub

'���ò˵��͹��߰�ť�Ŀ�������
Private Sub SetEnable()
    Dim bln�ѷ��� As Boolean
    Dim rstemp As New ADODB.Recordset
    
    With vsfList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
            
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
            If mnuEditReceive.Visible Then
                mnuEditReceive.Enabled = False
                tlbTool.Buttons("Receive").Enabled = False
            End If
            If mnuEditDisReceive.Visible Then
                mnuEditDisReceive.Enabled = False
                tlbTool.Buttons("DisReceive").Enabled = False
            End If
         Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            If .TextMatrix(.Row, .Cols - 4) = "" Then    'δ��˵�
                bln�ѷ��� = (vsfList.TextMatrix(vsfList.Row, 12) <> "")
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                If mnuEditReceive.Visible Then
                    mnuEditReceive.Enabled = bln�ѷ���
                    tlbTool.Buttons("Receive").Enabled = bln�ѷ���
                End If
                
                mnuEditDisReceive.Enabled = False
                tlbTool.Buttons("DisReceive").Enabled = False
                
                '�����������δ��ˣ�������ɾ��
                If mint�������� = 1 Then
                    If Val(.TextMatrix(.Row, .Cols - 3)) Mod 3 = 2 Then
                        mnuEditModify.Enabled = False
                        tlbTool.Buttons("Modify").Enabled = False
                        mnuEditReceive.Enabled = False
                        tlbTool.Buttons("Receive").Enabled = False
                        mnuEditDisReceive.Enabled = False
                        tlbTool.Buttons("DisReceive").Enabled = False
                        
                        mnuEditDel.Enabled = True
                        tlbTool.Buttons("Delete").Enabled = True
                    End If
                Else
                    If mnuEditDisReceive.Visible Then
                        If bln�ѷ��� Then
                            mnuEditDisReceive.Enabled = Not bln�ѷ���
                            tlbTool.Buttons("DisReceive").Enabled = Not bln�ѷ���
                        Else
                            mnuEditDisReceive.Enabled = False
                            tlbTool.Buttons("DisReceive").Enabled = False
                        End If
                    End If
                End If
                        
            ElseIf .TextMatrix(.Row, .Cols - 3) = 1 Then    '��˵�
                '�ж��Ƿ���ܣ���֧���ѳ������ݵĽ��ܹ��ܣ�����ȫ�˻��为���ķ�ʽ�������ΪҪʵ��������ܣ���Ҫ����ͳ��ʣ��������
                
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                If mnuEditReceive.Visible Then
                    mnuEditReceive.Enabled = False
                    tlbTool.Buttons("Receive").Enabled = False
                End If
                If mnuEditDisReceive.Visible Then
                    mnuEditDisReceive.Enabled = True
                    tlbTool.Buttons("DisReceive").Enabled = True
                End If
            Else   '2,3 ������
                If .TextMatrix(.Row, .Cols - 3) Mod 3 = 0 Then
                    .ToolTipText = "�������ݵ�ԭ����"
                    If mnuEditDisReceive.Visible = True Then
                        mnuEditDisReceive.Enabled = True
                        tlbTool.Buttons("DisReceive").Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, .Cols - 3) Mod 3 = 2 Then
                    .ToolTipText = "��������"
                    If mnuEditDisReceive.Visible = True Then
                        mnuEditDisReceive.Enabled = False
                        tlbTool.Buttons("DisReceive").Enabled = False
                    End If
                End If
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                If mnuEditReceive.Visible Then
                    mnuEditReceive.Enabled = False
                    tlbTool.Buttons("Receive").Enabled = False
                End If
            End If
        End If
        
    End With
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    Else
        strRange = "�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = MStrCaption
        
    objRow.Add "ʱ�䣺" & strRange
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡ����:" & Format(zlDataBase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfList
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.index = 1 Then
        mnuEditAddHank_Click
    ElseIf ButtonMenu.index = 2 Then
        mnuEditAddAuto_Click
    ElseIf ButtonMenu.index = 3 Then
        mnuEditAddAutoBySale_Click
    ElseIf ButtonMenu.index = 4 Then
        mnuEditAddMerge_Click
    End If
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub
'Ѱ����ĳһ����ȵ���
Public Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function

Private Sub subExcel(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String
    
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = MStrCaption
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "NO")))
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "�Ƴ��ⷿ��" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "��ҩ�ⷿ"))
    objRow.Add "����ⷿ��" & gstrDeptName
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "ժҪ:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "ժҪ"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "������")) & "  ��������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "��������"))
    
    objRow.Add "�����:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�����")) & "  �������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�������"))
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Function Is����(ByVal StrBillNo As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '�ȼ���ǲ������쵥
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(��ҩ��ʽ,0) ���� From ҩƷ�շ���¼ " & _
              " Where ����=6 And NO=[1] And ���ϵ�� = -1 and rownum = 1"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "����ǲ������쵥", StrBillNo)
    
    Is���� = Not (rsCheck!���� = 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub CheckNumber()
    '�����д������ʵ��������һ�£����ú�ɫ�����עʵ��������������
    Dim intRow As Integer
    Dim blnColor As Boolean

    With vsfDetail
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
            blnColor = False
            If .TextMatrix(intRow, .ColIndex("ҩƷID")) = "" Then Exit Sub
            If Val(.TextMatrix(intRow, .ColIndex("��д����"))) <> Val(.TextMatrix(intRow, .ColIndex("ʵ������"))) Then blnColor = True
            .Cell(flexcpForeColor, intRow, .ColIndex("ʵ������"), intRow, .ColIndex("ʵ������")) = IIf(blnColor, vbRed, vbBlack)
        Next
    End With
                
End Sub

