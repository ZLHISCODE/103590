VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmExternalAllocation 
   Caption         =   "������������"
   ClientHeight    =   9960
   ClientLeft      =   165
   ClientTop       =   870
   ClientWidth     =   14505
   Icon            =   "frmExternalAllocation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   14505
   StartUpPosition =   1  '����������
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   73
      Top             =   9585
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmExternalAllocation.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20505
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
   Begin VB.PictureBox picEdit 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   5040
      ScaleHeight     =   8655
      ScaleWidth      =   9255
      TabIndex        =   3
      Top             =   120
      Width           =   9255
      Begin VB.Frame fra������ϢZLBH 
         Caption         =   " ������Ϣ "
         Height          =   735
         Left            =   1080
         TabIndex        =   45
         Top             =   3600
         Width           =   7695
         Begin VB.TextBox txtZLBH��ַ 
            Height          =   270
            Left            =   1035
            MaxLength       =   250
            TabIndex        =   47
            Top             =   315
            Width           =   5415
         End
         Begin VB.Label lblZLBH��ַ 
            AutoSize        =   -1  'True
            Caption         =   "ZLBH��ַ"
            Height          =   180
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame fra������ϢEXE 
         Caption         =   " ������Ϣ "
         Height          =   735
         Left            =   1080
         TabIndex        =   35
         Top             =   2880
         Width           =   7695
         Begin VB.CommandButton cmd����·�� 
            Caption         =   "��"
            Height          =   270
            Left            =   6480
            TabIndex        =   38
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "��*��ѡ����"
            Top             =   315
            Width           =   270
         End
         Begin VB.TextBox txt�������·�� 
            Height          =   270
            Left            =   1400
            MaxLength       =   250
            TabIndex        =   37
            Top             =   315
            Width           =   5055
         End
         Begin VB.Label lbl�������·�� 
            AutoSize        =   -1  'True
            Caption         =   "�������·��"
            Height          =   180
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   1080
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   51
         Top             =   5280
         Width           =   7695
         _cx             =   13573
         _cy             =   1296
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmExternalAllocation.frx":0E1C
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
         OwnerDraw       =   1
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   5280
         Width           =   7695
         _cx             =   13573
         _cy             =   1296
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmExternalAllocation.frx":0EAD
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
         OwnerDraw       =   1
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
      Begin VB.Frame fraӦ�ó��� 
         Caption         =   " Ӧ�ó��� "
         Height          =   1815
         Left            =   120
         TabIndex        =   53
         Top             =   6360
         Width           =   7695
         Begin VB.PictureBox picСͼ�� 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   256
            Left            =   4627
            ScaleHeight     =   174.545
            ScaleMode       =   0  'User
            ScaleWidth      =   182.857
            TabIndex        =   61
            Top             =   682
            Width           =   256
            Begin VB.Image imgСͼ�� 
               Height          =   64
               Left            =   0
               Top             =   0
               Width           =   62
            End
         End
         Begin VB.CommandButton cmd��ͼ�� 
            Caption         =   "��"
            Height          =   240
            Left            =   5037
            TabIndex        =   70
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "��*��ѡ����"
            Top             =   1410
            Width           =   255
         End
         Begin VB.CommandButton cmdСͼ�� 
            Caption         =   "��"
            Height          =   240
            Left            =   5037
            TabIndex        =   62
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "��*��ѡ����"
            Top             =   690
            Width           =   255
         End
         Begin VB.CommandButton cmd��մ�ͼ�� 
            Caption         =   "��"
            Height          =   240
            Left            =   5342
            TabIndex        =   71
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "��*��ѡ����"
            Top             =   1410
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmd���Сͼ�� 
            Caption         =   "��"
            Height          =   240
            Left            =   5342
            TabIndex        =   63
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "��*��ѡ����"
            Top             =   690
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox pic��ͼ�� 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   350
            Left            =   4627
            ScaleHeight     =   360
            ScaleMode       =   0  'User
            ScaleWidth      =   360
            TabIndex        =   69
            Top             =   1355
            Width           =   360
            Begin VB.Image img��ͼ�� 
               Height          =   44
               Left            =   0
               Top             =   0
               Width           =   46
            End
         End
         Begin VB.ComboBox cbo������ 
            Height          =   300
            Left            =   1000
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1020
            Width           =   2055
         End
         Begin VB.CheckBox chk����ҽ������վ 
            Caption         =   "����ҽ������վ"
            Height          =   180
            Left            =   1000
            TabIndex        =   55
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkסԺҽ������վ 
            Caption         =   "סԺҽ������վ"
            Height          =   180
            Left            =   3012
            TabIndex        =   56
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkסԺ��ʿ����վ 
            Caption         =   "סԺ��ʿ����վ"
            Height          =   180
            Left            =   4905
            TabIndex        =   57
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cbo�˵� 
            Height          =   300
            Left            =   1000
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   660
            Width           =   2055
         End
         Begin VB.ComboBox cbo�Ҽ��˵� 
            Height          =   300
            Left            =   1000
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   1380
            Width           =   2055
         End
         Begin VB.Label lbl��ʾСͼ�� 
            AutoSize        =   -1  'True
            Caption         =   "Сͼ��"
            Height          =   180
            Left            =   4050
            TabIndex        =   60
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lbl��ʾ��ͼ�� 
            AutoSize        =   -1  'True
            Caption         =   "��ͼ��"
            Height          =   180
            Left            =   4050
            TabIndex        =   68
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label lbl�Ҽ��˵� 
            AutoSize        =   -1  'True
            Caption         =   "�Ҽ��˵�"
            Height          =   180
            Left            =   240
            TabIndex        =   66
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   420
            TabIndex        =   64
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lbl�˵� 
            AutoSize        =   -1  'True
            Caption         =   "�˵�"
            Height          =   180
            Left            =   600
            TabIndex        =   58
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblӦ�ó��� 
            AutoSize        =   -1  'True
            Caption         =   "Ӧ�ó���"
            Height          =   180
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame fra������ϢFTP 
         Caption         =   " ������Ϣ "
         Height          =   2175
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   7695
         Begin VB.TextBox txtFTP����Ŀ¼ 
            Height          =   270
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   28
            Top             =   1035
            Width           =   5175
         End
         Begin VB.CommandButton cmdFTP���Ӳ��� 
            Caption         =   "FTP���Ӳ���"
            Height          =   350
            Left            =   5025
            TabIndex        =   34
            Top             =   1710
            Width           =   1335
         End
         Begin VB.TextBox txtFTP��ַ 
            Height          =   270
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   20
            Top             =   315
            Width           =   2070
         End
         Begin VB.TextBox txtFTP���� 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   4305
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   26
            Top             =   675
            Width           =   2055
         End
         Begin VB.TextBox txtFTP�û��� 
            Height          =   270
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   24
            Top             =   675
            Width           =   2055
         End
         Begin VB.TextBox txtFTP����Ŀ¼ 
            Height          =   270
            Left            =   1200
            MaxLength       =   250
            TabIndex        =   30
            Top             =   1395
            Width           =   5175
         End
         Begin VB.TextBox txtFTP�˿� 
            Height          =   270
            Left            =   4305
            MaxLength       =   10
            TabIndex        =   22
            Top             =   315
            Width           =   2055
         End
         Begin VB.TextBox txt�ļ������� 
            Height          =   270
            Left            =   1185
            MaxLength       =   50
            TabIndex        =   33
            Top             =   1755
            Width           =   2055
         End
         Begin VB.CommandButton cmd����Ŀ¼ 
            Caption         =   "��"
            Height          =   270
            Left            =   6405
            TabIndex        =   31
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "��*��ѡ����"
            Top             =   1395
            Width           =   270
         End
         Begin VB.Label lblFTP����Ŀ¼ 
            AutoSize        =   -1  'True
            Caption         =   "FTP����Ŀ¼"
            Height          =   180
            Left            =   180
            TabIndex        =   27
            Top             =   1080
            Width           =   990
         End
         Begin VB.Label lbl�ļ������� 
            AutoSize        =   -1  'True
            Caption         =   "�ļ�������"
            Height          =   180
            Left            =   240
            TabIndex        =   32
            Top             =   1800
            Width           =   900
         End
         Begin VB.Label lblFTP�˿� 
            AutoSize        =   -1  'True
            Caption         =   "FTP�˿�"
            Height          =   180
            Left            =   3630
            TabIndex        =   21
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblFTP����Ŀ¼ 
            AutoSize        =   -1  'True
            Caption         =   "FTP����Ŀ¼"
            Height          =   180
            Left            =   180
            TabIndex        =   29
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label lblFTP���� 
            AutoSize        =   -1  'True
            Caption         =   "FTP����"
            Height          =   180
            Left            =   3630
            TabIndex        =   25
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lblFTP�û��� 
            AutoSize        =   -1  'True
            Caption         =   "FTP�û���"
            Height          =   180
            Left            =   360
            TabIndex        =   23
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblFTP��ַ 
            AutoSize        =   -1  'True
            Caption         =   "FTP��ַ"
            Height          =   180
            Left            =   540
            TabIndex        =   19
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.Frame fra���뷽ʽ 
         Caption         =   " ���뷽ʽ "
         Height          =   650
         Left            =   120
         TabIndex        =   13
         Top             =   1915
         Width           =   7695
         Begin VB.OptionButton opt���뷽ʽ 
            Caption         =   "ZLBH"
            Height          =   180
            Index           =   3
            Left            =   6120
            TabIndex        =   17
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton opt���뷽ʽ 
            Caption         =   "FTP"
            Height          =   180
            Index           =   2
            Left            =   4160
            TabIndex        =   16
            Top             =   300
            Width           =   615
         End
         Begin VB.OptionButton opt���뷽ʽ 
            Caption         =   "EXE"
            Height          =   180
            Index           =   1
            Left            =   2200
            TabIndex        =   15
            Top             =   300
            Width           =   615
         End
         Begin VB.OptionButton opt���뷽ʽ 
            Caption         =   "URL"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   300
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Frame fra�ӿڻ�����Ϣ 
         Caption         =   " �ӿڻ�����Ϣ "
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   7695
         Begin VB.ComboBox cbo�ӿ���� 
            Height          =   300
            Left            =   4250
            TabIndex        =   8
            Top             =   300
            Width           =   2655
         End
         Begin VB.TextBox txt��� 
            Height          =   270
            Left            =   645
            MaxLength       =   5
            TabIndex        =   6
            Top             =   315
            Width           =   2655
         End
         Begin VB.TextBox txt���� 
            Height          =   270
            Left            =   645
            MaxLength       =   50
            TabIndex        =   10
            Top             =   675
            Width           =   6255
         End
         Begin VB.TextBox txt˵�� 
            Height          =   630
            Left            =   645
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   1080
            Width           =   6255
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lbl˵�� 
            AutoSize        =   -1  'True
            Caption         =   "˵��"
            Height          =   180
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl�ӿ���� 
            AutoSize        =   -1  'True
            Caption         =   "�ӿ����"
            Height          =   180
            Left            =   3480
            TabIndex        =   7
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            Caption         =   "���"
            Height          =   180
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   360
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   5280
         Width           =   7695
         _cx             =   13573
         _cy             =   1296
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmExternalAllocation.frx":0F3E
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
         OwnerDraw       =   1
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
      Begin VB.Frame fra������ϢURL 
         Caption         =   " ������Ϣ "
         Height          =   1095
         Left            =   1080
         TabIndex        =   39
         Top             =   3240
         Width           =   7695
         Begin VB.OptionButton opt��������� 
            Caption         =   "Chrome"
            Height          =   180
            Index           =   1
            Left            =   2535
            TabIndex        =   42
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton opt��������� 
            Caption         =   "IE"
            Height          =   180
            Index           =   0
            Left            =   1320
            TabIndex        =   41
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.TextBox txtURL��ַ 
            Height          =   270
            Left            =   1320
            MaxLength       =   250
            TabIndex        =   44
            Top             =   675
            Width           =   5535
         End
         Begin VB.Label lblURL��ַ 
            AutoSize        =   -1  'True
            Caption         =   "URL��ַ"
            Height          =   180
            Left            =   510
            TabIndex        =   43
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lbl��������� 
            AutoSize        =   -1  'True
            Caption         =   "���������"
            Height          =   180
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Label lbl�б�˵�� 
         AutoSize        =   -1  'True
         Caption         =   "�б�˵��..."
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   120
         TabIndex        =   52
         Top             =   6120
         Width           =   990
      End
      Begin VB.Label lbl��ʾ��Ϣ 
         AutoSize        =   -1  'True
         Caption         =   "��ʾ��Ϣ..."
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   120
         TabIndex        =   48
         Top             =   4920
         Width           =   990
      End
      Begin VB.Label lblͼ����ʾ 
         AutoSize        =   -1  'True
         Caption         =   "˵������ͼ��Ҫ��24*24��ico��ʽ��Сͼ��Ҫ��16*16��ico��ʽ���㼤���ѡ��ʹ��"
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   120
         TabIndex        =   72
         Top             =   8280
         Width           =   6660
      End
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4575
      ScaleWidth      =   3105
      TabIndex        =   0
      Top             =   1080
      Width           =   3105
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   3570
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2475
         _Version        =   589884
         _ExtentX        =   4366
         _ExtentY        =   6297
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1800
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExternalAllocation.frx":0FCF
            Key             =   "ʡ��"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
      WordWrap        =   -1  'True
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
   Begin MSComDlg.CommonDialog cdl��Ƭ 
      Left            =   240
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.ImageManager imgFunc 
      Left            =   0
      Top             =   480
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmExternalAllocation.frx":32FD
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmExternalAllocation.frx":40D7
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmExternalAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const conPane_List = 201
Private Const conPane_Edit = 202
Private Const INTERNET_OPEN_TYPE_DIRECT     As Long = &H1           'direct to net

Private Const INTERNET_SERVICE_FTP          As Long = &H1
Private Const INTERNET_FLAG_KEEP_CONNECTION  As Long = &H400000    ' use keep-alive semantics
Private Const INTERNET_FLAG_PASSIVE         As Long = &H8000000   ' used for FTP connections

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_EDITBOX = &H10
Private Const BIF_USENEWUI = BIF_NEWDIALOGSTYLE Or BIF_EDITBOX
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'���ܣ�������Internet�ĻỰ
'˵����
'    sAgent--Ҫ����Internet�Ի���Ӧ�ó�����
'    lAccessType--�����������ʵ�����
'��ע�����lAccessType����ΪINTERNET_OPEN_TYPE_PRECONFIG������ʱ��Ҫ����
'    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings
'    ע���·���µ�ע�����ֵProxyEnable��ProxyServer�� ProxyOverride
'    sProxyName--ָ����������������֣�������������ΪINTERNET_OPEN_TYPE_PROXY����Ч
'    sProxyBypass--ָ����������������ֻ��ַ�������ô���ʱlpszProxyNameָ���Ľ�ʧЧ
'��������ֵ�������������ʧ�ܣ�lngINet Ϊ0��

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'���ܣ�����Internet���ӣ���FTP�Ự
'˵����
'    hInternetSession--����InternetOpen���ص�Internet�Ự���
'    sServerName--Ҫ���ӵķ����������ƻ�IP
'    nServerPort--Ҫ���ӵ�Internet�˿�
'    sUsername--��¼���û��ʺ�
'    sPassword--��¼�Ŀ���
'    lService--Ҫ���ӵķ��������ͣ�����������FTP�����������ӵ�����Ϊ����INTERNET_SERVICE_FTP��
'    lFlags--�������x8000000�����ӽ�ʹ�ñ���FTP���壬����0ʹ�÷Ǳ�������
'    lContext--��ʹ�ûص�����ʱʹ�øò�������ʹ�ûص����񴫵�0
'��������ֵ�������������ʧ�ܣ�lngINetConn Ϊ0

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'���ܣ��ر�Internet����


Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�
Private mintEditType As Integer    '��ǰ�༭��״̬��0-�鿴;1-����;2-�޸�
Private mlngDelID As Long          'ɾ����ID���޸�ͨ����ɾ�ٲ�ķ�ʽ
Private mbln��ʾͣ�� As Boolean

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'���뷽ʽ
Private Enum mTnterfaceType
    URL
    EXE
    FTP
    ZLBH
End Enum

'�������
Private Enum mBrowserType
    IE
    Chrome
End Enum

'�����
Private Enum mREPORT_COLUMN
    COL_ID
    col_�Ƿ�ͣ��
    COL_���
    col_���
    col_����
    col_���뷽ʽ
    col_˵��
End Enum

Private Sub InitCommandBars()
    '���ܣ���ʼ���˵�������ȫ���˵����������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar

    'CommandBars��������
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    Me.cbsThis.VisualTheme = xtpThemeOffice2003

    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With

    Me.cbsThis.EnableCustomization False
    Set cbsThis.Icons = zlcommfun.GetPubIcons
    '-----------------------------------------------------

    '�˵�����
    '-----------------------------------------------------
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagStretched)

    '***�ļ�
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        '�ļ�-Ԥ��
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        '�ļ�-��ӡ
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")

        '�ļ�-�˳�
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControlMain.BeginGroup = True
    End With

    '***�༭
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        '�༭-����
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        '�༭-�޸�
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        '�༭-ɾ��
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")

        '�༭-����
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Reuse, "����(&U)")
        '�༭-ͣ��
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Pause, "ͣ��(&P)")
    End With

    '***�鿴
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        '�鿴-������
        Set cbrControlMain = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        cbrControl.Checked = True

        '�鿴-״̬��
        Set cbrControlMain = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        cbrControlMain.Checked = True
        
        '�鿴-�б�
        Set cbrControlMain = .Add(xtpControlPopup, conMenu_View_Append, "�б�(&L)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, comMenu_LIS_ShowListHead, "��ʾ��ͣ��(&S)", -1, False)
        cbrControl.Checked = True
        mbln��ʾͣ�� = True
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
        cbrControlMain.BeginGroup = True
    End With

    '***����
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        '����-��������
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")

        '����-WEB
        Set cbrControlMain = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False

        '����-����
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
        cbrControlMain.BeginGroup = True
    End With
    '-----------------------------------------------------

    '�����
    '-----------------------------------------------------
    With Me.cbsThis.KeyBindings
        '��ӡ
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        'ɾ��
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
        '����
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        '�޸�
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        '��������
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    '-----------------------------------------------------

    '����������
    '-----------------------------------------------------
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched

    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_NewItem, "����")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")

        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Reuse, "����")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Pause, "ͣ��")

        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        cbrControlMain.BeginGroup = True
    End With

    '��ʾ���
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
    '-----------------------------------------------------
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlControl.RPTCopyToVSF(Me.rptList, Me.vfgList) Is Nothing Then Exit Sub
     
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "�������������嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub Load�ӿ����()
    '���ܣ�������б����еĽӿ��������������
    Dim rptRow As ReportRow
    Dim strTemp As String
    
    cbo�ӿ����.Clear
    
    If Me.rptList.Rows.Count > 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If InStr(";" & strTemp & ";", rptRow.Record(col_���).Value) < 1 Then
                    cbo�ӿ����.AddItem rptRow.Record(col_���).Value
                    strTemp = strTemp & IIf(strTemp = "", "", ";") & rptRow.Record(col_���).Value
                End If
            End If
        Next
        
        cbo�ӿ����.ListIndex = -1
    End If

End Sub

Private Sub cbo�ӿ����_GotFocus()
    zlControl.TxtSelAll cbo�ӿ����
End Sub

Private Sub cbo�ӿ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'-+_!@#$%^&*(){}[];:,.<>?/|\����������������������������%����&����", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    
    Select Case Control.ID
    Case conMenu_File_Preview
        'Ԥ��
        Call zlRptPrint(0)
        
    Case conMenu_File_Print
        '��ӡ
        Call zlRptPrint(1)
        
    Case conMenu_File_Exit
        '�˳�
        Unload Me
        
    Case conMenu_Edit_NewItem
        '����
        mintEditType = 1
        cmd��մ�ͼ��.Visible = False
        cmd���Сͼ��.Visible = False
        
        Call EnabledControl(mintEditType)
        Call ResetControl
        
    Case conMenu_Edit_Modify
        '�޸�
        mintEditType = 2
        mlngDelID = rptList.FocusedRow.Record.Item(mREPORT_COLUMN.COL_ID).Value
        Call EnabledControl(mintEditType)

    Case conMenu_Edit_Delete
        'ɾ��
        Call DeleteItem
        
    Case conMenu_Edit_Reuse
        '����
        Call StopAndStart(0)
        
    Case conMenu_Edit_Pause
        'ͣ��
        Call StopAndStart(1)
    
    Case conMenu_Edit_Save
        '����
        Call Save
        
    Case conMenu_Edit_Untread
        'ȡ��
        Call Untread
        
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case comMenu_LIS_ShowListHead
        mbln��ʾͣ�� = Not mbln��ʾͣ��
        Call RefreshList
    Case conMenu_View_Refresh
        Call RefreshList
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    End Select
End Sub

Private Sub StopAndStart(ByVal intMode As Integer)
    '���ܣ���Ŀͣ��
    'intMode��0-������1-ֹͣ
    Dim lngId As Long
    Dim strSql As String
    
    On Error GoTo ErrHandle
    
    lngId = rptList.FocusedRow.Record.Item(mREPORT_COLUMN.COL_ID).Value
    
    strSql = "Zl_��������Ŀ¼_Stop("
    'ID
    strSql = strSql & lngId
    '�Ƿ�ͣ��
    strSql = strSql & "," & intMode
    strSql = strSql & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call RefreshList(lngId)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DeleteItem()
    '���ܣ�ɾ��ѡ�е���Ŀ
    Dim lngId As Long
    Dim strSql As String
    Dim strMsg As String
    
    On Error GoTo ErrHandle
       
    strMsg = "���ɾ������Ŀ��"
    strMsg = strMsg & vbCrLf & "����" & rptList.FocusedRow.Record(mREPORT_COLUMN.col_����).Value
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
    lngId = rptList.FocusedRow.Record.Item(mREPORT_COLUMN.COL_ID).Value
    
    strSql = "Zl_��������Ŀ¼_Delete("
    'ID
    strSql = strSql & lngId
    strSql = strSql & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call RefreshList
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        '����
        Control.Enabled = (mintEditType = 0 And (zlStr.IsHavePrivs(mstrPrivs, "��ɾ��")))
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        '�޸�,ɾ��
        If mintEditType <> 0 Or rptList.FocusedRow Is Nothing Or rptList.FocusedRow.GroupRow Then
            Control.Enabled = False
        Else
            Control.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "��ɾ��"))
        End If
    Case conMenu_Edit_Reuse
        '����
        If mintEditType <> 0 Or rptList.FocusedRow Is Nothing Or rptList.FocusedRow.GroupRow Then
            Control.Enabled = False
        Else
            Control.Enabled = (rptList.FocusedRow.Record.Item(mREPORT_COLUMN.col_�Ƿ�ͣ��).Value = 1 And (zlStr.IsHavePrivs(mstrPrivs, "��ɾ��")))
        End If
    Case conMenu_Edit_Pause
        'ͣ��
        If mintEditType <> 0 Or rptList.FocusedRow Is Nothing Or rptList.FocusedRow.GroupRow Then
            Control.Enabled = False
        Else
            Control.Enabled = (rptList.FocusedRow.Record.Item(mREPORT_COLUMN.col_�Ƿ�ͣ��).Value = 0 And (zlStr.IsHavePrivs(mstrPrivs, "��ɾ��")))
        End If
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        '���棬ȡ��
        Control.Enabled = (mintEditType <> 0)
    Case conMenu_File_Preview, conMenu_File_Print
        'Ԥ������ӡ
        Control.Enabled = (mintEditType = 0)
    
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case comMenu_LIS_ShowListHead
        Control.Checked = mbln��ʾͣ��
    Case conMenu_View_Refresh
        Control.Enabled = (mintEditType = 0)
    End Select
End Sub

Private Sub cmdFTP���Ӳ���_Click()
    Dim lngINet As Long
    Dim lngINetConn As Long
    
    lngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If lngINet <= 0 Then
        MsgBox "��������ʧ�ܣ�", vbExclamation, "FTP���Ӳ���"
        Exit Sub
    Else
        lngINetConn = InternetConnect(lngINet, txtFTP��ַ.Text, Val(txtFTP�˿�.Text), txtFTP�û���.Text, txtFTP����.Text, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_PASSIVE, 0)
        If lngINetConn = 0 Then
            Call InternetCloseHandle(lngINet)
            MsgBox "��������ʧ�ܣ�", vbExclamation, "FTP���Ӳ���"
            Exit Sub
        Else
            Call InternetCloseHandle(lngINet)
        End If
    End If

    MsgBox "�������ӳɹ���", vbInformation, "FTP���Ӳ���"
End Sub

Private Sub Save()
    Dim rsTemp As ADODB.Recordset
    Dim lngId As Long       '����ID
    Dim strSql As String
    Dim str��ַ As String
    Dim int���뷽ʽ As Integer
    Dim date��ҩʱ�� As Date
    Dim blnInTrans As Boolean
    Dim arrSql As Variant
    Dim i As Integer
    Dim intSel As Integer
    
    On Error GoTo ErrHandle
    
    date��ҩʱ�� = sys.Currentdate
    arrSql = Array()
    
    '�����Ƿ�¼������
    '--------------------------
    If Trim(txt���.Text) = "" Then
        MsgBox "���δ¼�룡"
        Exit Sub
    End If
    
    If cbo�ӿ����.Text = "" Then
        MsgBox "�ӿ����δ¼�룡"
        Exit Sub
    End If
    
    If Trim(txt����.Text) = "" Then
        MsgBox "����δ¼�룡"
        Exit Sub
    End If
    
    If opt���뷽ʽ(mTnterfaceType.URL).Value Then
        '***URL
        
        If InStr(txtURL��ַ.Text, "://") < 2 Then
            MsgBox "URL��ַ��ʽ����ȷ��"
            Exit Sub
        End If
        
    ElseIf opt���뷽ʽ(mTnterfaceType.EXE).Value Then
        '***EXE
        
        If InStr(txt�������·��.Text, ".exe") < 2 Then
            MsgBox "�������·����ʽ����ȷ��"
            Exit Sub
        End If
        
    ElseIf opt���뷽ʽ(mTnterfaceType.FTP).Value Then
        '***FTP
        
        If Trim(txtFTP��ַ.Text) = "" Then
            MsgBox "FTP��ַΪ�գ�"
            Exit Sub
        End If
        
        If Trim(txtFTP�û���.Text) = "" Then
            MsgBox "FTP�û���Ϊ�գ�"
            Exit Sub
        End If
        
        If Trim(txtFTP�˿�.Text) = "" Then
            MsgBox "FTP�˿�Ϊ�գ�"
            Exit Sub
        End If
    ElseIf opt���뷽ʽ(mTnterfaceType.ZLBH).Value Then
        '***ZLBH
        
        If Trim(txtZLBH��ַ.Text) = "" Then
            MsgBox "ZLBH��ַΪ�գ�"
            Exit Sub
        End If
    End If
    
    '--------------------------
    
    '�޸�ʱ��ɾ��
    If mintEditType = 2 Then
        gstrSql = "Zl_��������Ŀ¼_Delete("
        'ID
        gstrSql = gstrSql & mlngDelID
        gstrSql = gstrSql & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSql
    End If
    
    '��ȡ����·����ID
    '--------------------------
    strSql = "Select ��������Ŀ¼_Id.Nextval ID From Dual"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    lngId = rsTemp!ID
    '--------------------------
    
    '���ݲ�ͬ�Ľ��뷽ʽ��������ϲ�ͬ������
    '--------------------------
    If opt���뷽ʽ(mTnterfaceType.URL).Value Then
        intSel = mTnterfaceType.URL
        int���뷽ʽ = 1
        str��ַ = "'" & txtURL��ַ.Text & "'"
    ElseIf opt���뷽ʽ(mTnterfaceType.EXE).Value Then
        intSel = mTnterfaceType.EXE
        int���뷽ʽ = 2
        str��ַ = "'" & txt�������·��.Text & "'"
    ElseIf opt���뷽ʽ(mTnterfaceType.FTP).Value Then
        intSel = mTnterfaceType.FTP
        int���뷽ʽ = 3
        str��ַ = "Null"
    ElseIf opt���뷽ʽ(mTnterfaceType.ZLBH).Value Then
        int���뷽ʽ = 4
        str��ַ = "'" & txtZLBH��ַ.Text & "'"
    End If
    '--------------------------
    
    
    gstrSql = "Zl_��������Ŀ¼_Insert("
    'ID
    gstrSql = gstrSql & lngId
    '���
    gstrSql = gstrSql & "," & Val(txt���.Text)
    '���
    gstrSql = gstrSql & ",'" & cbo�ӿ����.Text & "'"
    '����
    gstrSql = gstrSql & ",'" & txt����.Text & "'"
    '˵��
    gstrSql = gstrSql & ",'" & txt˵��.Text & "'"
    '���뷽ʽ
    gstrSql = gstrSql & "," & int���뷽ʽ
    '���������
    gstrSql = gstrSql & "," & IIf(opt���뷽ʽ(mTnterfaceType.URL).Value, IIf(opt���������(mBrowserType.IE).Value, 1, 2), "Null")
    'Ӧ�ó���
    gstrSql = gstrSql & ",'" & chk����ҽ������վ.Value & chkסԺҽ������վ.Value & chkסԺ��ʿ����վ.Value & "'"
    '��ַ
    gstrSql = gstrSql & "," & str��ַ
    '�Ƿ�ͣ��
    gstrSql = gstrSql & ",0"
    'Ftp��ַ
    gstrSql = gstrSql & ",'" & IIf(opt���뷽ʽ(mTnterfaceType.FTP).Value, txtFTP��ַ.Text, "") & "'"
    'Ftp����Ŀ¼
    gstrSql = gstrSql & ",'" & IIf(opt���뷽ʽ(mTnterfaceType.FTP).Value, txtFTP����Ŀ¼.Text, "") & "'"
    'Ftp�û���
    gstrSql = gstrSql & ",'" & IIf(opt���뷽ʽ(mTnterfaceType.FTP).Value, txtFTP�û���.Text, "") & "'"
    'Ftp����
    gstrSql = gstrSql & ",'" & IIf(opt���뷽ʽ(mTnterfaceType.FTP).Value, zlStr.Sm4EncryptEcb(txtFTP����.Text), "") & "'"
    'Ftp����Ŀ¼
    gstrSql = gstrSql & ",'" & IIf(opt���뷽ʽ(mTnterfaceType.FTP).Value, txtFTP����Ŀ¼.Text, "") & "'"
    'Ftp�˿�
    gstrSql = gstrSql & ",'" & IIf(opt���뷽ʽ(mTnterfaceType.FTP).Value, txtFTP�˿�.Text, "") & "'"
    'Ftp�ļ���
    gstrSql = gstrSql & ",'" & IIf(opt���뷽ʽ(mTnterfaceType.FTP).Value, txt�ļ�������.Text, "") & "'"
    '�˵���ʾ
    gstrSql = gstrSql & "," & cbo�˵�.ListIndex
    '��������ʾ
    gstrSql = gstrSql & "," & cbo������.ListIndex
    '�Ҽ��˵���ʾ
    gstrSql = gstrSql & "," & cbo�Ҽ��˵�.ListIndex
    '�޸���
    gstrSql = gstrSql & ",'" & gstrUserName & "'"
    '�޸�ʱ��
    gstrSql = gstrSql & ",to_date('" & date��ҩʱ�� & "','yyyy-MM-dd hh24:mi:ss')"
    gstrSql = gstrSql & ")"
    
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSql
    
    If Not opt���뷽ʽ(mTnterfaceType.ZLBH).Value Then
        
        
        With vsfList(intSel)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("���")) <> "" And .TextMatrix(i, .ColIndex("����ֵ")) <> "" Then
                    gstrSql = "Zl_�������ò���_Insert("
                    '�ӿ�id
                    gstrSql = gstrSql & lngId
                    '���
                    gstrSql = gstrSql & "," & .TextMatrix(i, .ColIndex("���"))
                    '����ֵ
                    gstrSql = gstrSql & ",'" & .TextMatrix(i, .ColIndex("����ֵ")) & "'"
                    '��ע
                    gstrSql = gstrSql & ",'" & .TextMatrix(i, .ColIndex("��ע")) & "'"
                    'Sql
                    gstrSql = gstrSql & ",'" & .TextMatrix(i, .ColIndex("����Դ")) & "'"
                    gstrSql = gstrSql & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSql
                End If
            Next
        End With
    End If
    
    '���д���ҩ����
    '--------------------------
    gcnOracle.BeginTrans
    blnInTrans = True

    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "��������Ŀ¼")
    Next
    
    gcnOracle.CommitTrans
    blnInTrans = False
    '--------------------------
    
    '����ͼ��
    '--------------------------
    Call sys.SaveLob(100, 31, lngId, imgСͼ��.Tag)
    Call sys.SaveLob(100, 32, lngId, img��ͼ��.Tag)
    '--------------------------

    Call RefreshList(lngId)
    
    mintEditType = 0
    Call EnabledControl(mintEditType)
    
    Exit Sub
ErrHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd��ͼ��_Click()
    Dim pic As stdole.StdPicture
    Dim lngH As Long
    Dim lngW As Long
                    
    With cdl��Ƭ
        .CancelError = True
        .Filter = "ͼƬ�ļ�(*.ico)|*.ico"
        
        On Error Resume Next
        .ShowOpen
        
        If err <> 0 Then
            'ûѡ���ļ�
            err.Clear
        Else
            '�ж�ͼ��Ĵ�С�ߴ��Ƿ����Ҫ�󣬴�ͼ��Ϊ24*24����
            '----------------------------------------------
            Set pic = LoadPicture(.FileName)
            
            lngH = Int(pic.Height * 0.567 / 15 + 0.5)
            lngW = Int(pic.Width * 0.567 / 15 + 0.5)
            
            If lngH <> 24 Or lngW <> 24 Then
                MsgBox "��ѡ������Ϊ24*24��ͼ�꣡", vbInformation, gstrSysName
                Exit Sub
            End If
            '----------------------------------------------
            
            img��ͼ��.Picture = LoadPicture(.FileName)

            If err <> 0 Then
                MsgBox "ͼƬ�ļ���Ч�����ļ������ڡ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            img��ͼ��.Tag = .FileName
            
            cmd��մ�ͼ��.Visible = True
        End If
    End With
End Sub

Private Sub cmd��մ�ͼ��_Click()
    Set img��ͼ��.Picture = Nothing
    img��ͼ��.Tag = ""
    cmd��մ�ͼ��.Visible = False
End Sub

Private Sub cmd���Сͼ��_Click()
    Set imgСͼ��.Picture = Nothing
    imgСͼ��.Tag = ""
    cmd���Сͼ��.Visible = False
End Sub

Private Sub Untread()
    mintEditType = 0
    
    With rptList
        If Me.rptList.Rows.Count > 0 Then
            Call RefreshList
        End If
    End With
    
    Call EnabledControl(mintEditType)
End Sub

Public Function BrowseForFolder(Optional sTitle As String = "��ѡ���ļ���") As String
    Dim intNull As Integer, lngIDList As Long
    Dim strPath As String, udtBI As BrowseInfo

    With udtBI
        .hWndOwner = 0 ' Me.hWnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
    End With
    
    lngIDList = SHBrowseForFolder(udtBI)
    
    If lngIDList Then
        strPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lngIDList, strPath
        CoTaskMemFree lngIDList
        intNull = InStr(strPath, vbNullChar)
        
        If intNull Then
          strPath = Left$(strPath, intNull - 1)
        End If
    End If

    BrowseForFolder = strPath
End Function

Private Sub cmd����Ŀ¼_Click()
    '��������Ŀ¼·��
    
    txtFTP����Ŀ¼.Text = BrowseForFolder
End Sub

Private Sub cmdСͼ��_Click()
    Dim pic As stdole.StdPicture
    Dim lngH As Long
    Dim lngW As Long

    With cdl��Ƭ
        .CancelError = True
        .Filter = "ͼƬ�ļ�(*.ico)|*.ico"
        
        On Error Resume Next
        .ShowOpen
        
        If err <> 0 Then
            'ûѡ���ļ�
            err.Clear
        Else
            '�ж�ͼ��Ĵ�С�ߴ��Ƿ����Ҫ��Сͼ��Ϊ16*16����
            '----------------------------------------------
            Set pic = LoadPicture(.FileName)
            
            lngH = Int(pic.Height * 0.567 / 15 + 0.5)
            lngW = Int(pic.Width * 0.567 / 15 + 0.5)
            
            If lngH <> 16 Or lngW <> 16 Then
                MsgBox "��ѡ������Ϊ16*16��ͼ�꣡", vbInformation, gstrSysName
                Exit Sub
            End If
            '----------------------------------------------

            imgСͼ��.Picture = LoadPicture(.FileName)
            
            Debug.Print imgСͼ��.Picture.Width
            
            If err <> 0 Then
                MsgBox "ͼƬ�ļ���Ч�����ļ������ڡ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            imgСͼ��.Tag = .FileName
            
            cmd���Сͼ��.Visible = True
        End If
    End With
End Sub

Private Sub cmd����·��_Click()
    '���س������·��
    Dim str������ As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln���� As Boolean
    Dim i As Integer
    
    str������ = ""
    bln���� = False
    
    If txt�������·��.Text <> "" Then
        If InStr(txt�������·��.Text, ".exe[") > 0 Then
            str������ = Mid(txt�������·��.Text, InStr(txt�������·��.Text, ".exe[") + 4)
            
            '�ȶԵ�ǰ�Ĳ����б���л���
            If vsfList(mTnterfaceType.EXE).Rows > 1 Then
                With rsTemp
                    If .State = 1 Then .Close
                    
                    .Fields.Append "���", adDouble, 18, adFldIsNullable
                    .Fields.Append "����ֵ", adLongVarChar, 200, adFldIsNullable
                    .Fields.Append "��ע", adLongVarChar, 500, adFldIsNullable
                    .Fields.Append "����Դ", adLongVarChar, 2000, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .Open
                    
                    bln���� = True
                    
                    For i = 1 To vsfList(mTnterfaceType.EXE).Rows - 1
                        .AddNew
                        
                        !��� = vsfList(mTnterfaceType.EXE).TextMatrix(i, vsfList(mTnterfaceType.EXE).ColIndex("���"))
                        !����ֵ = vsfList(mTnterfaceType.EXE).TextMatrix(i, vsfList(mTnterfaceType.EXE).ColIndex("����ֵ"))
                        !��ע = vsfList(mTnterfaceType.EXE).TextMatrix(i, vsfList(mTnterfaceType.EXE).ColIndex("��ע"))
                        !����Դ = vsfList(mTnterfaceType.EXE).TextMatrix(i, vsfList(mTnterfaceType.EXE).ColIndex("����Դ"))
                        
                        .Update
                    Next
                End With
            End If
        End If
    End If
    
    With cdl��Ƭ
        .CancelError = True
        .Filter = "��ִ���ļ�(*.exe)|*.exe"
        
        On Error Resume Next
        .ShowOpen
        
        If err <> 0 Then
            'ûѡ���ļ�
            err.Clear
        Else
            If err <> 0 Then
                MsgBox "�����ļ���Ч�����ļ������ڡ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            txt�������·��.Text = .FileName & str������
            
            If bln���� Then
                rsTemp.Filter = ""
                With vsfList(mTnterfaceType.EXE)
                    .Redraw = flexRDNone
                    
                    For i = 1 To rsTemp.RecordCount
                        .TextMatrix(i, .ColIndex("���")) = rsTemp!���
                        .TextMatrix(i, .ColIndex("����ֵ")) = rsTemp!����ֵ
                        .TextMatrix(i, .ColIndex("��ע")) = zlcommfun.NVL(rsTemp!��ע, "")
                        .TextMatrix(i, .ColIndex("����Դ")) = zlcommfun.NVL(rsTemp!Sqltext, "")
                        
                        rsTemp.MoveNext
                    Next
                        
                    .Redraw = flexRDDirect
                End With
            End If
        End If
    End With
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = picList.hwnd
    Case conPane_Edit
        Item.Handle = picEdit.hwnd
    End Select
End Sub

Private Sub InitPanes()
    '���ܣ���ʼ��DockingPane�ؼ�
    Dim panList As Pane
    Dim panEdit As Pane

    Set panList = dkpMan.CreatePane(conPane_List, 500, 1000, DockLeftOf, Nothing)
    panList.Title = "��Ϣ�б�"
    panList.Options = PaneNoCaption
    
    Set panEdit = dkpMan.CreatePane(conPane_Edit, 500, 1000, DockRightOf, panList)
    panEdit.Title = "��Ϣ�༭"
    panEdit.Options = PaneNoCaption

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
End Sub

Private Sub InitReportControl()
    '���ܣ���ʼ��ReportControl�ؼ�
    Dim objCol As ReportColumn

    With rptList
        .Columns.DeleteAll
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)

        Set objCol = .Columns.Add(COL_ID, "ID", 0, False)
        objCol.Sortable = False
        objCol.Visible = False
        
        Set objCol = .Columns.Add(col_�Ƿ�ͣ��, "�Ƿ�ͣ��", 0, False)
        objCol.Sortable = False
        objCol.Visible = False
        
        Set objCol = .Columns.Add(COL_���, "���", 150, True)
        objCol.Sortable = True
        objCol.Visible = True
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentRight
        .SortOrder.Add objCol       'Ĭ����������
        
        Set objCol = .Columns.Add(col_���, "���", 200, True)
        objCol.Sortable = False
        objCol.Visible = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.TreeColumn = True

        Set objCol = .Columns.Add(col_����, "����", 400, True)
        objCol.Sortable = True
        objCol.Visible = True
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft

        Set objCol = .Columns.Add(col_���뷽ʽ, "���뷽ʽ", 200, True)
        objCol.Sortable = True
        objCol.Visible = True
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft

        Set objCol = .Columns.Add(col_˵��, "˵��", 500, True)
        objCol.Sortable = False
        objCol.Visible = True
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft

        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = (objCol.Index = col_���)
        Next
        
        .AllowColumnRemove = False
        .MultipleSelection = False  '���������ѡ�񡣻�����SelectionChanged�¼�
        .ShowItemsInGroups = False  '����ʾ�ѷ������

        .GroupsOrder.Add .Columns(col_���)

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ������..."
        End With
    End With
End Sub

Private Sub InitControl()
    '���ܣ���ʼ���ؼ������Լ�����
    
    With cbo�˵�
        .Clear
        .AddItem "0-����ʾ"
        .AddItem "1-��ʾ���Ӳ˵���"
        .AddItem "2-��ʾ�����˵���"
        .ListIndex = 0
    End With
    
    With cbo������
        .Clear
        .AddItem "0-����ʾ"
        .AddItem "1-��ʾ���ӹ�������"
        .AddItem "2-��ʾ������������"
        .ListIndex = 0
    End With
    
    With cbo�Ҽ��˵�
        .Clear
        .AddItem "0-����ʾ"
        .AddItem "1-��ʾ���Ӳ˵���"
        .AddItem "2-��ʾ�����˵���"
        .ListIndex = 0
    End With
    
    With img��ͼ��
        .Left = pic��ͼ��.ScaleLeft
        .Top = pic��ͼ��.ScaleTop
        .Width = pic��ͼ��.ScaleWidth
        .Height = pic��ͼ��.ScaleHeight
    End With
    
    With imgСͼ��
        .Left = picСͼ��.ScaleLeft
        .Top = picСͼ��.ScaleTop
        .Width = picСͼ��.ScaleWidth
        .Height = picСͼ��.ScaleHeight
    End With
End Sub

Private Sub InitvsfList()
    '���ܣ���ʼ�������б�
    
    With vsfList(mTnterfaceType.URL)
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("����ֵ")) = "����ID|����ID|ҽ��ID|����ID|��¼�û���|����Ա���|����Ա����|����Դ��ȡ|"
    End With
    
    With vsfList(mTnterfaceType.EXE)
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("����ֵ")) = "����ID|����ID|ҽ��ID|����ID|��¼�û���|����Ա���|����Ա����|����Դ��ȡ|"
    End With
    
    With vsfList(mTnterfaceType.FTP)
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("����ֵ")) = "����ID|����ID|ҽ��ID|����ID|��¼�û���|����Ա���|����Ա����|����Դ��ȡ|"
    End With
    
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    
    mintEditType = 0
    
    Call zlcommfun.SetWindowsInTaskBar(Me.hwnd, False)
    Call InitCommandBars
    Call InitPanes
    Call InitReportControl
    Call InitControl
    Call InitvsfList

    Call RefreshList
    Call EnabledControl(mintEditType)
    
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub EnabledControl(ByVal intEditType As Integer)
    '���ܣ����Ʊ༭����Ƿ����
    
    picList.Enabled = (intEditType = 0)
    
    '�ı���
    txt���.Enabled = (intEditType <> 0)
    cbo�ӿ����.Enabled = (intEditType <> 0)
    txt����.Enabled = (intEditType <> 0)
    txt˵��.Enabled = (intEditType <> 0)
    
    txt�������·��.Enabled = (intEditType <> 0)
    cmd����·��.Enabled = (intEditType <> 0)
    
    txtURL��ַ.Enabled = (intEditType <> 0)
    
    txtFTP��ַ.Enabled = (intEditType <> 0)
    txtFTP����Ŀ¼.Enabled = (intEditType <> 0)
    txtFTP�û���.Enabled = (intEditType <> 0)
    txtFTP����.Enabled = (intEditType <> 0)
    txtFTP�˿�.Enabled = (intEditType <> 0)
    txt�ļ�������.Enabled = (intEditType <> 0)
    txtFTP����Ŀ¼.Enabled = (intEditType <> 0)
    cmd����Ŀ¼.Enabled = (intEditType <> 0)
    
    txtZLBH��ַ.Enabled = (intEditType <> 0)
    
    cbo�˵�.Enabled = (intEditType <> 0)
    cbo������.Enabled = (intEditType <> 0)
    cbo�Ҽ��˵�.Enabled = (intEditType <> 0)
    
    cmdСͼ��.Enabled = (intEditType <> 0)
    cmd��ͼ��.Enabled = (intEditType <> 0)
    
    cmd���Сͼ��.Enabled = (intEditType <> 0)
    cmd��մ�ͼ��.Enabled = (intEditType <> 0)
    
    cmdFTP���Ӳ���.Enabled = (intEditType <> 0)
    
    chk����ҽ������վ.Enabled = (intEditType <> 0)
    chkסԺҽ������վ.Enabled = (intEditType <> 0)
    chkסԺ��ʿ����վ.Enabled = (intEditType <> 0)
    
    opt���뷽ʽ(mTnterfaceType.EXE).Enabled = (intEditType <> 0)
    opt���뷽ʽ(mTnterfaceType.FTP).Enabled = (intEditType <> 0)
    opt���뷽ʽ(mTnterfaceType.URL).Enabled = (intEditType <> 0)
    opt���뷽ʽ(mTnterfaceType.ZLBH).Enabled = (intEditType <> 0)
    
    vsfList(mTnterfaceType.EXE).Enabled = (intEditType <> 0)
    vsfList(mTnterfaceType.FTP).Enabled = (intEditType <> 0)
    vsfList(mTnterfaceType.URL).Enabled = (intEditType <> 0)
    
    opt���������(mBrowserType.IE).Enabled = (intEditType <> 0)
    opt���������(mBrowserType.Chrome).Enabled = (intEditType <> 0)
End Sub

Private Sub ResetControl()
    '���ܣ���ձ༭��������пռ������
    Dim objTemp As Control
    
    For Each objTemp In Me.Controls
        '����ı�
        If TypeName(objTemp) = "TextBox" Then
            objTemp.Text = ""
        End If
        
        '���ö�ѡ��
        If TypeName(objTemp) = "CheckBox" Then
            objTemp.Value = 1
        End If
    Next
    
    '���õ�ѡ��
    opt���뷽ʽ(mTnterfaceType.URL).Value = True
    opt���������(mBrowserType.IE).Value = True
    
    '���ñ��
    vsfList(mTnterfaceType.URL).Rows = 1
    vsfList(mTnterfaceType.EXE).Rows = 1
    vsfList(mTnterfaceType.FTP).Rows = 1
    
    '���ͼ��
    Set imgСͼ��.Picture = Nothing
    imgСͼ��.Tag = ""
    
    Set img��ͼ��.Picture = Nothing
    img��ͼ��.Tag = ""
    
    '���������б�
    cbo�˵�.ListIndex = 0
    cbo������.ListIndex = 0
    cbo�Ҽ��˵�.ListIndex = 0

End Sub

Private Sub RefreshList(Optional ByVal lngPart As Long)
    '���ܣ�ˢ���б�
    Dim rsData As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim ObjItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    On Error GoTo ErrHandle
    
    'Select
    gstrSql = "Select a.Id, a.�Ƿ�ͣ��, a.���, a.���, a.����, a.˵��, Decode(a.���뷽ʽ, 1, 'URL', 2, 'EXE', 3, 'FTP', 4, 'ZLBH',a.���뷽ʽ) As ���뷽ʽ "

    'From
    gstrSql = gstrSql & " From ��������Ŀ¼ A "
    
    'Where
    If Not mbln��ʾͣ�� Then
        gstrSql = gstrSql & " Where a.�Ƿ�ͣ�� = 0 "
    End If
    
    'Order By
    gstrSql = gstrSql & " Order By a.��� "
    
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    '�б��������
    '---------------------------------------
    rptList.Records.DeleteAll
    
    Do While Not rsData.EOF
        Set objRecord = rptList.Records.Add()
        
        Set ObjItem = objRecord.AddItem(Val(rsData!ID))
        Set ObjItem = objRecord.AddItem(Val(rsData!�Ƿ�ͣ��))
        
        Set ObjItem = objRecord.AddItem(String(5 - Len(CStr(rsData!���)), " ") & CStr(rsData!���))    '�ո���λ���������ִ�С˳������
        ObjItem.ForeColor = IIf(Val(rsData!�Ƿ�ͣ��) = 0, vbBlack, vbRed)
        
        Set ObjItem = objRecord.AddItem(CStr(rsData!���))
        ObjItem.ForeColor = IIf(Val(rsData!�Ƿ�ͣ��) = 0, vbBlack, vbRed)
        
        Set ObjItem = objRecord.AddItem(CStr(rsData!����))
        ObjItem.ForeColor = IIf(Val(rsData!�Ƿ�ͣ��) = 0, vbBlack, vbRed)
        
        Set ObjItem = objRecord.AddItem(CStr(rsData!���뷽ʽ))
        ObjItem.ForeColor = IIf(Val(rsData!�Ƿ�ͣ��) = 0, vbBlack, vbRed)
        
        Set ObjItem = objRecord.AddItem(CStr("" & rsData!˵��))
        ObjItem.ForeColor = IIf(Val(rsData!�Ƿ�ͣ��) = 0, vbBlack, vbRed)
        
        rsData.MoveNext
    Loop
    
    rptList.Populate
    '---------------------------------------
    
    If lngPart <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(COL_ID).Value) = lngPart Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    
    '״̬����ʾ
    '---------------------------------------
    Me.stbThis.Panels(2).Text = "����" & rsData.RecordCount & "��ӿ�"
    '---------------------------------------
    
    Call rptList_SelectionChanged
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub opt���뷽ʽ_Click(Index As Integer)
    Call DynamicArrange
End Sub

Private Sub picList_Resize()
    err = 0: On Error Resume Next
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
End Sub

Private Sub picEdit_Resize()
    err = 0: On Error Resume Next
    
    '�༭�̶�����
    '-------------------------------------
    With fra�ӿڻ�����Ϣ
        .Top = 50
        .Left = 100
        .Width = picEdit.Width - .Left - 100
    End With
    
    With fra���뷽ʽ
        .Top = fra�ӿڻ�����Ϣ.Top + fra�ӿڻ�����Ϣ.Height + 100
        .Left = fra�ӿڻ�����Ϣ.Left
        .Width = picEdit.Width - .Left - 100
    End With
    '-------------------------------------
    
    '��̬����
    Call DynamicArrange
End Sub

Private Sub DynamicArrange()
    '���ܣ�����ѡ��Ľ������Ͷ�̬���пؼ�
    
    err = 0: On Error Resume Next
    
    '��̬����
    '-------------------------------------
    '***URL
    With fra������ϢURL
        .Visible = (opt���뷽ʽ(mTnterfaceType.URL).Value)
        .Top = fra���뷽ʽ.Top + fra���뷽ʽ.Height + 100
        .Left = fra���뷽ʽ.Left
        .Width = picEdit.Width - .Left - 100
    End With
    
    '***EXE
    With fra������ϢEXE
        .Visible = (opt���뷽ʽ(mTnterfaceType.EXE).Value)
        .Top = fra���뷽ʽ.Top + fra���뷽ʽ.Height + 100
        .Left = fra���뷽ʽ.Left
        .Width = picEdit.Width - .Left - 100
    End With
    
    '***FTP
    With fra������ϢFTP
        .Visible = (opt���뷽ʽ(mTnterfaceType.FTP).Value)
        .Top = fra���뷽ʽ.Top + fra���뷽ʽ.Height + 100
        .Left = fra���뷽ʽ.Left
        .Width = picEdit.Width - .Left - 100
    End With
    
    '***ZLBH
    With fra������ϢZLBH
        .Visible = (opt���뷽ʽ(mTnterfaceType.ZLBH).Value)
        .Top = fra���뷽ʽ.Top + fra���뷽ʽ.Height + 100
        .Left = fra���뷽ʽ.Left
        .Width = picEdit.Width - .Left - 100
    End With
    
    '-------------------------------------

    '��β��̬����
    '-------------------------------------
    With lbl��ʾ��Ϣ
        .Left = fra���뷽ʽ.Left
        .Visible = True
        If opt���뷽ʽ(mTnterfaceType.URL).Value Then
            '--URL
            .Top = fra������ϢURL.Top + fra������ϢURL.Height + 50
            .Caption = "˵����URL��ַ�еġ���������[1]��[2]��[3]�ȱ�ʾ���磺http://192.168.1.4:8055/All/ResultDetail.aspx?MOD=UIS&&ID=[1]�����ţ�2016073074�����������봫�Ρ�" & vbCrLf & _
                    "ϵͳ�̶���������У�����ID������ID������ID��ҽ��ID����¼�û���������Ա��ţ�����Ա������"
        ElseIf opt���뷽ʽ(mTnterfaceType.EXE).Value Then
            '--EXE
            .Top = fra������ϢEXE.Top + fra������ϢEXE.Height + 50
            .Caption = "˵����EXE������õġ���������[1]��[2]��[3]�ȱ�ʾ���磺c\appsoft\zlhis+.exe[USER]/[����]/[���Ӵ�]��" & vbCrLf & _
                    "ϵͳ�̶���������У�����ID������ID������ID��ҽ��ID����¼�û���������Ա��ţ�����Ա������"
        ElseIf opt���뷽ʽ(mTnterfaceType.FTP).Value Then
            '--FTP
            .Top = fra������ϢFTP.Top + fra������ϢFTP.Height + 50
            .Caption = "˵�����ļ����Բ���[1]�ķ�ʽ���á�ϵͳ�̶���������У�����ID������ID������ID��ҽ��ID����¼�û���������Ա��ţ�����Ա�����ȣ�����������ṩ������Դ��ȡ�ķ�ʽ����ϴ������ֵ��"
        ElseIf opt���뷽ʽ(mTnterfaceType.ZLBH).Value Then
            .Visible = False
        End If
        
        .ToolTipText = .Caption
    End With
    
    '�ڶ�̬��������б�߶�ǰ��ȷ��ʣ�¿ռ�ĸ߶�
    lbl�б�˵��.Caption = "˵��������ֵΪ����Դ��ȡʱ��SQL�漰�����̶�����[����ID]��[����ID]��[����ID]��[ҽ��ID]" & vbCrLf & _
                "����ID�����ﲡ�˴���ֵΪ����ID��סԺ���˴���ֵΪ��ҳID"
    
    If ((picEdit.Height - stbThis.Height) - (lbl��ʾ��Ϣ.Top + lbl��ʾ��Ϣ.Height) - lblͼ����ʾ.Height - fraӦ�ó���.Height - lbl�б�˵��.Height - 350 > vsfList(mTnterfaceType.URL).RowHeightMin * 7) Or (opt���뷽ʽ(mTnterfaceType.ZLBH).Value) Then
        '˳������
        
        With vsfList(mTnterfaceType.URL)
            .Visible = (opt���뷽ʽ(mTnterfaceType.URL).Value)
            .Top = lbl��ʾ��Ϣ.Top + lbl��ʾ��Ϣ.Height + 100
            .Left = fra���뷽ʽ.Left
            .Width = picEdit.Width - .Left - 100
            .Height = vsfList(mTnterfaceType.URL).RowHeightMin * 7
        End With
        
        With vsfList(mTnterfaceType.EXE)
            .Visible = (opt���뷽ʽ(mTnterfaceType.EXE).Value)
            .Top = lbl��ʾ��Ϣ.Top + lbl��ʾ��Ϣ.Height + 100
            .Left = fra���뷽ʽ.Left
            .Width = picEdit.Width - .Left - 100
            .Height = vsfList(mTnterfaceType.URL).RowHeightMin * 7
        End With
        
        With vsfList(mTnterfaceType.FTP)
            .Visible = (opt���뷽ʽ(mTnterfaceType.FTP).Value)
            .Top = lbl��ʾ��Ϣ.Top + lbl��ʾ��Ϣ.Height + 100
            .Left = fra���뷽ʽ.Left
            .Width = picEdit.Width - .Left - 100
            .Height = vsfList(mTnterfaceType.URL).RowHeightMin * 7
        End With
        
        With lbl�б�˵��
            .Visible = Not (opt���뷽ʽ(mTnterfaceType.ZLBH).Value)
            .Top = vsfList(mTnterfaceType.URL).Top + vsfList(mTnterfaceType.URL).Height + 100
            .Left = fra���뷽ʽ.Left
            .ToolTipText = .Caption
        End With
        
        With fraӦ�ó���
            If (opt���뷽ʽ(mTnterfaceType.ZLBH).Value) Then
                .Top = fra������ϢZLBH.Top + fra������ϢZLBH.Height + 100
            Else
                .Top = lbl�б�˵��.Top + lbl�б�˵��.Height + 100
            End If
            
            .Left = fra���뷽ʽ.Left
            .Width = picEdit.Width - .Left - 100
        End With
        
        With lblͼ����ʾ
            .Top = fraӦ�ó���.Top + fraӦ�ó���.Height + 50
            .Left = fra���뷽ʽ.Left
            .ToolTipText = .Caption
        End With
    Else
        '���´��µ��Ϸ���top
        With lblͼ����ʾ
            .Top = picEdit.Height - stbThis.Height - .Height - 100
            .Left = fra���뷽ʽ.Left
            .ToolTipText = .Caption
        End With
        
        With fraӦ�ó���
            .Top = lblͼ����ʾ.Top - .Height - 50
            .Left = fra���뷽ʽ.Left
            .Width = picEdit.Width - .Left - 100
        End With
    
        With lbl�б�˵��
            .Visible = True
            .Top = fraӦ�ó���.Top - .Height - 100
            .Left = fra���뷽ʽ.Left
            .ToolTipText = .Caption
        End With
        
        '����洰�����������
        With vsfList(mTnterfaceType.URL)
            .Visible = (opt���뷽ʽ(mTnterfaceType.URL).Value)
            .Top = lbl��ʾ��Ϣ.Top + lbl��ʾ��Ϣ.Height + 100
            .Left = fra���뷽ʽ.Left
            .Width = picEdit.Width - .Left - 100
            .Height = lbl�б�˵��.Top - .Top - 50
        End With
        
        With vsfList(mTnterfaceType.EXE)
            .Visible = (opt���뷽ʽ(mTnterfaceType.EXE).Value)
            .Top = lbl��ʾ��Ϣ.Top + lbl��ʾ��Ϣ.Height + 100
            .Left = fra���뷽ʽ.Left
            .Width = picEdit.Width - .Left - 100
            .Height = lbl�б�˵��.Top - .Top - 50
        End With
        
        With vsfList(mTnterfaceType.FTP)
            .Visible = (opt���뷽ʽ(mTnterfaceType.FTP).Value)
            .Top = lbl��ʾ��Ϣ.Top + lbl��ʾ��Ϣ.Height + 100
            .Left = fra���뷽ʽ.Left
            .Width = picEdit.Width - .Left - 100
            .Height = lbl�б�˵��.Top - .Top - 50
        End With
    End If
    '-------------------------------------
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub RefreshInfo(ByVal lngId As Long)
    '���ܣ�����idˢ�µ�ǰ��ʾ����
    Dim rsData As ADODB.Recordset
    Dim strСͼ�� As String
    Dim str��ͼ�� As String
    Dim i As Integer

    On Error GoTo ErrHandle
    
    If lngId = 0 Then Exit Sub
    
    Call ResetControl
    
    '��ȡ��������
    '---------------------------------
    gstrSql = "Select a.���, a.���, a.����, a.˵��, a.���뷽ʽ, a.���������, a.Ӧ�ó���, a.��ַ, a.Ftp��ַ, a.Ftp����Ŀ¼,a.Ftp�û���, a.Ftp����, a.Ftp����Ŀ¼," & vbNewLine & _
            "       a.Ftp�˿�, a.Ftp�ļ���, a.�˵���ʾ, a.��������ʾ, a.�Ҽ��˵���ʾ" & vbNewLine & _
            "From ��������Ŀ¼ A" & vbNewLine & _
            "Where a.Id = [1]"

    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngId)
    
    Me.txt���.Text = rsData!���

    Call Load�ӿ����
    For i = 1 To cbo�ӿ����.ListCount
        If cbo�ӿ����.List(i - 1) = rsData!��� Then
            cbo�ӿ����.ListIndex = i - 1
            Exit For
        End If
    Next

    Me.txt����.Text = rsData!����
    Me.txt˵��.Text = zlcommfun.NVL(rsData!˵��, "")
    
    Me.txtFTP��ַ.Text = zlcommfun.NVL(rsData!Ftp��ַ, "")
    Me.txtFTP����Ŀ¼.Text = zlcommfun.NVL(rsData!Ftp����Ŀ¼, "")
    Me.txtFTP�û���.Text = zlcommfun.NVL(rsData!Ftp�û���, "")
    
    If IsNull(rsData!Ftp����) Then
        Me.txtFTP����.Text = ""
    Else
        '����
        Me.txtFTP����.Text = zlStr.Sm4DecryptEcb(rsData!Ftp����)
    End If
    
    Me.txtFTP����Ŀ¼.Text = zlcommfun.NVL(rsData!Ftp����Ŀ¼, "")
    Me.txtFTP�˿�.Text = zlcommfun.NVL(rsData!Ftp�˿�, "")
    Me.txt�ļ�������.Text = zlcommfun.NVL(rsData!Ftp�ļ���, "")
    
    Me.opt���뷽ʽ(Val(rsData!���뷽ʽ) - 1).Value = True
    
    If opt���뷽ʽ(mTnterfaceType.URL).Value Then
        Me.txtURL��ַ.Text = zlcommfun.NVL(rsData!��ַ, "")
    ElseIf opt���뷽ʽ(mTnterfaceType.EXE).Value Then
        Me.txt�������·��.Text = zlcommfun.NVL(rsData!��ַ, "")
    ElseIf opt���뷽ʽ(mTnterfaceType.ZLBH).Value Then
        Me.txtZLBH��ַ.Text = zlcommfun.NVL(rsData!��ַ, "")
    End If
    
    If zlcommfun.NVL(rsData!���������, "") = "" Then
        Me.opt���������(mBrowserType.IE).Value = True
    Else
        Me.opt���������(Val(rsData!���������) - 1).Value = True
    End If
    Me.chk����ҽ������վ.Value = Mid(rsData!Ӧ�ó���, 1, 1)
    Me.chkסԺҽ������վ.Value = Mid(rsData!Ӧ�ó���, 2, 1)
    Me.chkסԺ��ʿ����վ.Value = Mid(rsData!Ӧ�ó���, 3, 1)
    
    cbo�˵�.ListIndex = rsData!�˵���ʾ
    cbo������.ListIndex = rsData!��������ʾ
    cbo�Ҽ��˵�.ListIndex = rsData!�Ҽ��˵���ʾ
    '---------------------------------
    
    '��ȡ������
    '---------------------------------
    gstrSql = "Select a.���, a.����ֵ, a.��ע, a.Sqltext From �������ò��� A Where �ӿ�id = [1]"
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngId)
    
    If opt���뷽ʽ(mTnterfaceType.URL).Value Then
        With vsfList(mTnterfaceType.URL)
            .Rows = 1
            .Redraw = flexRDNone
            
            Do While Not rsData.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("���")) = rsData!���
                .TextMatrix(.Rows - 1, .ColIndex("����ֵ")) = rsData!����ֵ
                .TextMatrix(.Rows - 1, .ColIndex("��ע")) = zlcommfun.NVL(rsData!��ע, "")
                .TextMatrix(.Rows - 1, .ColIndex("����Դ")) = zlcommfun.NVL(rsData!Sqltext, "")
    
                rsData.MoveNext
            Loop
            
            .Redraw = flexRDDirect
            
        End With
    End If
    
    If opt���뷽ʽ(mTnterfaceType.EXE).Value Then
        With vsfList(mTnterfaceType.EXE)
            .Rows = 1
            .Redraw = flexRDNone
            
            Do While Not rsData.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("���")) = rsData!���
                .TextMatrix(.Rows - 1, .ColIndex("����ֵ")) = rsData!����ֵ
                .TextMatrix(.Rows - 1, .ColIndex("��ע")) = zlcommfun.NVL(rsData!��ע, "")
                .TextMatrix(.Rows - 1, .ColIndex("����Դ")) = zlcommfun.NVL(rsData!Sqltext, "")
    
                rsData.MoveNext
            Loop
            
            .Redraw = flexRDDirect
            
        End With
    End If
    
    If opt���뷽ʽ(mTnterfaceType.FTP).Value Then
        With vsfList(mTnterfaceType.FTP)
            .Rows = 1
            .Redraw = flexRDNone
            
            Do While Not rsData.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("���")) = rsData!���
                .TextMatrix(.Rows - 1, .ColIndex("����ֵ")) = rsData!����ֵ
                .TextMatrix(.Rows - 1, .ColIndex("��ע")) = zlcommfun.NVL(rsData!��ע, "")
                .TextMatrix(.Rows - 1, .ColIndex("����Դ")) = zlcommfun.NVL(rsData!Sqltext, "")
    
                rsData.MoveNext
            Loop
            
            .Redraw = flexRDDirect
            
        End With
    End If
    '---------------------------------
    
    '��ȡͼ������
    '---------------------------------
    strСͼ�� = sys.Readlob(100, 31, lngId)
    str��ͼ�� = sys.Readlob(100, 32, lngId)
    
    imgСͼ��.Picture = LoadPicture(strСͼ��)
    imgСͼ��.Tag = strСͼ��
    
    img��ͼ��.Picture = LoadPicture(str��ͼ��)
    img��ͼ��.Tag = str��ͼ��
    
    cmd���Сͼ��.Visible = (strСͼ�� <> "")
    cmd��մ�ͼ��.Visible = (str��ͼ�� <> "")
    '---------------------------------
    
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptList_SelectionChanged()
    Dim lngId As Long
    
    With rptList
        If .FocusedRow Is Nothing Then
            lngId = 0
        ElseIf .FocusedRow.GroupRow = True Then
            lngId = 0
        Else
            lngId = .FocusedRow.Record.Item(mREPORT_COLUMN.COL_ID).Value
        End If
        Call RefreshInfo(lngId)
    End With
End Sub

Private Sub AutoLoading(ByVal objText As TextBox, ByVal intIndex As Integer)
    '���ܣ�����URL��ַ��������·���Ĳ������Զ����ض�Ӧ������ŵ����
    '˵����������"[1]��[2]��....[n]"����ʽ����
    'objText���ı������
    'intIndex���������
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim str��� As String
    Dim strTemp As String
    Dim strExistPars As String
    Dim intCount As Integer
    Dim i As Integer
    Dim n As Integer

    'ͳ��"["�ĸ���
    intCount = (Len(objText.Text) - Len(Replace(objText.Text, "[", ""))) / Len("[")
    
    lngStart = 1
    lngEnd = 1
    
    For i = 1 To intCount
        lngStart = InStr(lngEnd, objText.Text, "[")
        If lngStart = 0 Then Exit For
        
        lngEnd = InStr(lngStart, objText.Text, "]")
        If lngEnd = 0 Then Exit For
        
        If lngStart + 1 < lngEnd Then       '[lngStart + 1 < lngEnd]��ʾ"[?]"�м����ٺ���һ���ַ�
            str��� = Mid(objText.Text, lngStart + 1, lngEnd - lngStart - 1)
            strTemp = str���
            
            '��֤�Ƿ�Ϊ������
            '---------------------------
            strTemp = strTemp & vbCr
            For n = 0 To 9
                strTemp = Replace(strTemp, CStr(n), "")
            Next
            
            If strTemp = vbCr Then
                '�ռ�������
                strExistPars = strExistPars & IIf(strExistPars = "", "", ",") & str���
            
                With vsfList(intIndex)
                    If .FindRow(str���, , .ColIndex("���")) < 1 Then
                        '��������
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, .ColIndex("���")) = str���
                    End If
                End With
            End If
            '---------------------------
        End If
    Next
    
    'ɾ�������ڵĲ���
    With vsfList(intIndex)
        If strExistPars = "" Then
            '�޲���ʱ��ձ��
            .Rows = 1
        Else
            '�в�������Ҫ���ȽϺ���ɾ��
            strTemp = ""
            
            For i = 1 To .Rows - 1
                If InStr("," & strExistPars & ",", "," & .TextMatrix(i, .ColIndex("���")) & ",") < 1 Then
                    '�ռ���ɾ���Ĳ�����
                    strTemp = strTemp & IIf(strTemp = "", "", ",") & .TextMatrix(i, .ColIndex("���"))
                End If
            Next
                             
            For i = 0 To UBound(Split(strTemp, ","))
                .RemoveItem .FindRow(Split(strTemp, ",")(i), , .ColIndex("���"))
            Next
        End If
    End With
    
End Sub

Private Sub txtFTP��ַ_GotFocus()
    zlControl.TxtSelAll txtFTP��ַ
End Sub

Private Sub txtFTP��ַ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789.", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFTP�˿�_GotFocus()
    zlControl.TxtSelAll txtFTP�˿�
End Sub

Private Sub txtFTP�˿�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFTP����Ŀ¼_GotFocus()
    zlControl.TxtSelAll txtFTP����Ŀ¼
End Sub

Private Sub txtFTP����Ŀ¼_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtFTP�û���_GotFocus()
    zlControl.TxtSelAll txtFTP�û���
End Sub

Private Sub txtFTP�û���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'-+_!@#$%^&*(){}[];:,.<>?/|\����������������������������%����&����", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtFTP����_GotFocus()
    zlControl.TxtSelAll txtFTP����
End Sub

Private Sub txtFTP����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtURL��ַ_Change()
    Call AutoLoading(txtURL��ַ, 0)
End Sub

Private Sub txtURL��ַ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtZLBH��ַ_GotFocus()
    zlControl.TxtSelAll txtZLBH��ַ
End Sub

Private Sub txtZLBH��ַ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���_GotFocus()
    zlControl.TxtSelAll txt���
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub

    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt�������·��_Change()
    Call AutoLoading(txt�������·��, 1)
End Sub

Private Sub txt�������·��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_GotFocus()
    zlControl.TxtSelAll txt˵��
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If InStr(1, "'-+_!@#$%^&*(){}[];:<>?/|\������������������������%����&����", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�ļ�������_Change()
    Call AutoLoading(txt�ļ�������, 2)
End Sub

Private Sub txt�ļ�������_GotFocus()
    zlControl.TxtSelAll txt�ļ�������
End Sub

Private Sub txt�ļ�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtFTP����Ŀ¼_GotFocus()
    zlControl.TxtSelAll txtFTP����Ŀ¼
End Sub

Private Sub txtFTP����Ŀ¼_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfList_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim str����Դ As String

    With vsfList(Index)
        If Col = .ColIndex("����ֵ") Then
            str����Դ = .TextMatrix(Row, .ColIndex("����Դ"))
            
            '������ֵ���ǡ�����Դ��ȡ��ʱ����Ҫ�������Դ�ֶ��е����ݡ�
            .TextMatrix(Row, .ColIndex("����Դ")) = ""

            Select Case .TextMatrix(Row, Col)
            Case "����Դ��ȡ"
                .TextMatrix(.Row, .ColIndex("����Դ")) = frmExternalAllocationData.ShowMe(Me, str����Դ)
            End Select
        End If
    End With
End Sub

Private Sub vsfList_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList(Index)
        If Col <> .ColIndex("���") Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfList_ChangeEdit(Index As Integer)
    With vsfList(Index)
        If .Col = .ColIndex("����ֵ") Then
            .Col = .ColIndex("��ע")
        End If
    End With
End Sub

Private Sub vsfList_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(1, "'-+_!@#$%^&*(){}[];:<>?/|\������������������������%����&����", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
