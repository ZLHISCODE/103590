VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmParClinic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ٴ���������"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   Icon            =   "frmParClinic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   12630
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   12
      Left            =   2340
      ScaleHeight     =   7425
      ScaleWidth      =   10185
      TabIndex        =   273
      Tag             =   "j"
      Top             =   0
      Width           =   10215
      Begin TabDlg.SSTab StabNurse 
         Height          =   7245
         Left            =   105
         TabIndex        =   274
         Top             =   90
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   12779
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "��¼��"
         TabPicture(0)   =   "frmParClinic.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "���µ� "
         TabPicture(1)   =   "frmParClinic.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lbl(22)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lbl(23)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lbl(24)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lbl(30)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lbl(31)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lbl(32)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lblBreathe"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "lbl(33)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "lbl(34)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "lbl(35)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Line2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "lbl(36)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "lbl(37)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "lbl(38)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "lbl(39)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "lbl(27)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Line1"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "lbl(40)"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "lbl(41)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "chk(127)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "picControl(1)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "ud(22)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "ud(23)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "ud(24)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "ud(26)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "ud(27)"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "fra(15)"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "txtUD(22)"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "txtUD(23)"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "txtUD(24)"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "txtUD(26)"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "txtUD(27)"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).Control(32)=   "cbo(18)"
         Tab(1).Control(32).Enabled=   0   'False
         Tab(1).Control(33)=   "cbo(17)"
         Tab(1).Control(33).Enabled=   0   'False
         Tab(1).Control(34)=   "picLineColor(2)"
         Tab(1).Control(34).Enabled=   0   'False
         Tab(1).Control(35)=   "cbo(15)"
         Tab(1).Control(35).Enabled=   0   'False
         Tab(1).Control(36)=   "cbo(16)"
         Tab(1).Control(36).Enabled=   0   'False
         Tab(1).Control(37)=   "txt(28)"
         Tab(1).Control(37).Enabled=   0   'False
         Tab(1).Control(38)=   "chk(121)"
         Tab(1).Control(38).Enabled=   0   'False
         Tab(1).Control(39)=   "chk(122)"
         Tab(1).Control(39).Enabled=   0   'False
         Tab(1).Control(40)=   "chk(123)"
         Tab(1).Control(40).Enabled=   0   'False
         Tab(1).Control(41)=   "chk(124)"
         Tab(1).Control(41).Enabled=   0   'False
         Tab(1).Control(42)=   "chk(125)"
         Tab(1).Control(42).Enabled=   0   'False
         Tab(1).Control(43)=   "chk(126)"
         Tab(1).Control(43).Enabled=   0   'False
         Tab(1).Control(44)=   "chk(128)"
         Tab(1).Control(44).Enabled=   0   'False
         Tab(1).Control(45)=   "Frame1(4)"
         Tab(1).Control(45).Enabled=   0   'False
         Tab(1).Control(46)=   "picPoly"
         Tab(1).Control(46).Enabled=   0   'False
         Tab(1).Control(47)=   "PicInsert"
         Tab(1).Control(47).Enabled=   0   'False
         Tab(1).Control(48)=   "chk(130)"
         Tab(1).Control(48).Enabled=   0   'False
         Tab(1).Control(49)=   "chk(131)"
         Tab(1).Control(49).Enabled=   0   'False
         Tab(1).Control(50)=   "picLineColor(3)"
         Tab(1).Control(50).Enabled=   0   'False
         Tab(1).Control(51)=   "cbo(19)"
         Tab(1).Control(51).Enabled=   0   'False
         Tab(1).Control(52)=   "txt(25)"
         Tab(1).Control(52).Enabled=   0   'False
         Tab(1).Control(53)=   "chk(129)"
         Tab(1).Control(53).Enabled=   0   'False
         Tab(1).Control(54)=   "picLineColor(1)"
         Tab(1).Control(54).Enabled=   0   'False
         Tab(1).Control(55)=   "picOut"
         Tab(1).Control(55).Enabled=   0   'False
         Tab(1).Control(56)=   "picEnemaStool"
         Tab(1).Control(56).Enabled=   0   'False
         Tab(1).Control(57)=   "chk(132)"
         Tab(1).Control(57).Enabled=   0   'False
         Tab(1).Control(58)=   "picLineColor(4)"
         Tab(1).Control(58).Enabled=   0   'False
         Tab(1).Control(59)=   "chk(180)"
         Tab(1).Control(59).Enabled=   0   'False
         Tab(1).ControlCount=   60
         TabCaption(2)   =   "����ͼ"
         TabPicture(2)   =   "frmParClinic.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lbl(11)"
         Tab(2).Control(1)=   "fra1(1)"
         Tab(2).Control(2)=   "fra1(0)"
         Tab(2).Control(3)=   "fra2"
         Tab(2).Control(4)=   "cbo(6)"
         Tab(2).Control(5)=   "chk(109)"
         Tab(2).Control(6)=   "chk(110)"
         Tab(2).Control(7)=   "chk(111)"
         Tab(2).Control(8)=   "fraPoliceLine"
         Tab(2).ControlCount=   9
         Begin VB.CheckBox chk 
            Caption         =   "������Ŀ������ʾ"
            Height          =   180
            Index           =   180
            Left            =   -74685
            TabIndex        =   339
            Top             =   5805
            Width           =   2055
         End
         Begin VB.PictureBox picLineColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   4
            Left            =   -72045
            ScaleHeight     =   165
            ScaleWidth      =   270
            TabIndex        =   324
            TabStop         =   0   'False
            Top             =   2760
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "���µ���ӡʱ,������·�������˵����Ϣ"
            Height          =   195
            Index           =   132
            Left            =   -74685
            TabIndex        =   336
            Top             =   4950
            Width           =   3840
         End
         Begin VB.PictureBox picEnemaStool 
            BorderStyle     =   0  'None
            Height          =   190
            Left            =   -73140
            ScaleHeight     =   195
            ScaleWidth      =   2565
            TabIndex        =   354
            TabStop         =   0   'False
            Top             =   6870
            Width           =   2565
            Begin VB.OptionButton OptEnemaStool 
               Caption         =   "���ӷ�ĸ"
               Height          =   180
               Index           =   1
               Left            =   1230
               TabIndex        =   356
               Top             =   5
               Width           =   1050
            End
            Begin VB.OptionButton OptEnemaStool 
               Caption         =   "���±�"
               Height          =   180
               Index           =   0
               Left            =   15
               TabIndex        =   355
               Top             =   5
               Value           =   -1  'True
               Width           =   960
            End
         End
         Begin VB.PictureBox picOut 
            BorderStyle     =   0  'None
            Height          =   190
            Left            =   -73140
            ScaleHeight     =   195
            ScaleWidth      =   2565
            TabIndex        =   350
            TabStop         =   0   'False
            Top             =   6615
            Width           =   2565
            Begin VB.OptionButton OptOut 
               Caption         =   "ҽ������"
               Height          =   180
               Index           =   1
               Left            =   1230
               TabIndex        =   352
               Top             =   5
               Width           =   1170
            End
            Begin VB.OptionButton OptOut 
               Caption         =   "��Ժ��ʽ"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   351
               Top             =   5
               Value           =   -1  'True
               Width           =   1170
            End
         End
         Begin VB.PictureBox picLineColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   1
            Left            =   -66525
            ScaleHeight     =   165
            ScaleWidth      =   270
            TabIndex        =   361
            TabStop         =   0   'False
            Top             =   1365
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������ע�������ٴ�����ʱ,ֹͣǰһ��������ע"
            Height          =   180
            Index           =   129
            Left            =   -74700
            TabIndex        =   330
            Top             =   3330
            Width           =   4500
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   25
            Left            =   -73740
            MaxLength       =   2
            TabIndex        =   326
            Text            =   "14"
            Top             =   3045
            Width           =   255
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   19
            ItemData        =   "frmParClinic.frx":68A6
            Left            =   -71175
            List            =   "frmParClinic.frx":68A8
            Style           =   2  'Dropdown List
            TabIndex        =   329
            Top             =   2985
            Width           =   1725
         End
         Begin VB.PictureBox picLineColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   3
            Left            =   -72765
            ScaleHeight     =   165
            ScaleWidth      =   270
            TabIndex        =   327
            TabStop         =   0   'False
            Top             =   3045
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "Ӥ�����µ�����������0��ʼ"
            Height          =   180
            Index           =   131
            Left            =   -68325
            TabIndex        =   359
            Top             =   1110
            Width           =   2595
         End
         Begin VB.CheckBox chk 
            Caption         =   "���µ����ʱ��ӡҽԺ����"
            Height          =   180
            Index           =   130
            Left            =   -68325
            TabIndex        =   358
            Top             =   840
            Width           =   2550
         End
         Begin VB.PictureBox PicInsert 
            BorderStyle     =   0  'None
            Height          =   190
            Left            =   -73140
            ScaleHeight     =   195
            ScaleWidth      =   2565
            TabIndex        =   346
            TabStop         =   0   'False
            Top             =   6360
            Width           =   2565
            Begin VB.OptionButton OptInsert 
               Caption         =   "����/����"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   347
               Top             =   5
               Value           =   -1  'True
               Width           =   1170
            End
            Begin VB.OptionButton OptInsert 
               Caption         =   "����/����"
               Height          =   180
               Index           =   1
               Left            =   1230
               TabIndex        =   348
               Top             =   5
               Width           =   1170
            End
         End
         Begin VB.PictureBox picPoly 
            BorderStyle     =   0  'None
            Height          =   190
            Left            =   -73140
            ScaleHeight     =   195
            ScaleWidth      =   3360
            TabIndex        =   341
            TabStop         =   0   'False
            Top             =   6105
            Width           =   3360
            Begin VB.OptionButton optPloy 
               Caption         =   "б��"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   342
               Top             =   5
               Value           =   -1  'True
               Width           =   720
            End
            Begin VB.OptionButton optPloy 
               Caption         =   "ֱ��"
               Height          =   180
               Index           =   1
               Left            =   1230
               TabIndex        =   343
               Top             =   5
               Width           =   720
            End
            Begin VB.OptionButton optPloy 
               Caption         =   "�����"
               Height          =   180
               Index           =   2
               Left            =   2400
               TabIndex        =   344
               Top             =   5
               Width           =   840
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "���µ��ļ��Ŀ�ʼʱ��"
            Height          =   660
            Index           =   4
            Left            =   -68325
            TabIndex        =   387
            Top             =   6390
            Width           =   2970
            Begin VB.OptionButton optFileTime 
               Caption         =   "���ʱ��"
               Height          =   195
               Index           =   1
               Left            =   1680
               TabIndex        =   389
               Top             =   300
               Width           =   1125
            End
            Begin VB.OptionButton optFileTime 
               Caption         =   "��Ժʱ��"
               Height          =   195
               Index           =   0
               Left            =   300
               TabIndex        =   388
               Top             =   300
               Value           =   -1  'True
               Width           =   1125
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "���µ�����ʾ���˵������Ϣ"
            Height          =   180
            Index           =   128
            Left            =   -74685
            TabIndex        =   338
            Top             =   5520
            Width           =   2790
         End
         Begin VB.CheckBox chk 
            Caption         =   "���µ���ӡʱ,�����������(�������ʵ���ʹ����Ч)"
            Height          =   180
            Index           =   126
            Left            =   -74685
            TabIndex        =   335
            Top             =   4680
            Width           =   4770
         End
         Begin VB.CheckBox chk 
            Caption         =   "���ܡ�������Ŀ��ʾ�������ݣ�������ʾ���죩"
            Height          =   180
            Index           =   125
            Left            =   -74685
            TabIndex        =   332
            Top             =   3870
            Width           =   4215
         End
         Begin VB.CheckBox chk 
            Caption         =   "ȫ���������¼�롢��ʾ����ʱ��(h)"
            Height          =   180
            Index           =   124
            Left            =   -74685
            TabIndex        =   333
            Top             =   4140
            Width           =   3330
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����������ע������Ժ,������ǵ�ҳ��ʾ��ȫ"
            Height          =   180
            Index           =   123
            Left            =   -74685
            TabIndex        =   331
            Top             =   3600
            Width           =   4500
         End
         Begin VB.CheckBox chk 
            Caption         =   "���µ�ֻ��ʾ��Ʊ�ʶʱ�����Զ�ת��Ϊ��Ժ"
            Height          =   180
            Index           =   122
            Left            =   -74685
            TabIndex        =   337
            Top             =   5235
            Width           =   4020
         End
         Begin VB.CheckBox chk 
            Caption         =   "���µ�������ÿҳ�������ڸ�ʽ�̶�Ϊ:��-��-��"
            Height          =   180
            Index           =   121
            Left            =   -74700
            TabIndex        =   334
            Top             =   4410
            Width           =   4365
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   28
            Left            =   -72840
            MaxLength       =   2
            TabIndex        =   323
            Text            =   "v"
            Top             =   2745
            Width           =   255
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   16
            ItemData        =   "frmParClinic.frx":68AA
            Left            =   -73215
            List            =   "frmParClinic.frx":68AC
            Style           =   2  'Dropdown List
            TabIndex        =   317
            Top             =   1800
            Width           =   1590
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   15
            ItemData        =   "frmParClinic.frx":68AE
            Left            =   -73215
            List            =   "frmParClinic.frx":68B0
            Style           =   2  'Dropdown List
            TabIndex        =   313
            Top             =   1470
            Width           =   1590
         End
         Begin VB.PictureBox picLineColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   2
            Left            =   -71070
            ScaleHeight     =   165
            ScaleWidth      =   270
            TabIndex        =   315
            TabStop         =   0   'False
            Top             =   1500
            Width           =   300
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   17
            ItemData        =   "frmParClinic.frx":68B2
            Left            =   -71490
            List            =   "frmParClinic.frx":68B4
            Style           =   2  'Dropdown List
            TabIndex        =   321
            Top             =   2445
            Width           =   2430
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   18
            ItemData        =   "frmParClinic.frx":68B6
            Left            =   -71490
            List            =   "frmParClinic.frx":68B8
            Style           =   2  'Dropdown List
            TabIndex        =   319
            Top             =   2115
            Width           =   2430
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   27
            Left            =   -73215
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   298
            Text            =   "0"
            Top             =   525
            Width           =   350
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   26
            Left            =   -72390
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   307
            Text            =   "0"
            Top             =   840
            Width           =   350
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   24
            Left            =   -70650
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   304
            Text            =   "6"
            Top             =   525
            Width           =   350
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   23
            Left            =   -72000
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   301
            Text            =   "18"
            Top             =   525
            Width           =   350
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   22
            Left            =   -72390
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   310
            Text            =   "0"
            Top             =   1155
            Width           =   350
         End
         Begin VB.Frame fra 
            Caption         =   "�����Զ���־--��ɫ    "
            Height          =   4845
            Index           =   15
            Left            =   -68325
            TabIndex        =   360
            Top             =   1395
            Width           =   2970
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   21
               ItemData        =   "frmParClinic.frx":68BA
               Left            =   1605
               List            =   "frmParClinic.frx":68BC
               Style           =   2  'Dropdown List
               TabIndex        =   386
               Top             =   4410
               Width           =   840
            End
            Begin VB.CheckBox chk 
               Caption         =   "��˳���ڵ�������"
               Height          =   180
               Index           =   120
               Left            =   135
               TabIndex        =   382
               Top             =   3630
               Width           =   2505
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   33
               ItemData        =   "frmParClinic.frx":68BE
               Left            =   705
               List            =   "frmParClinic.frx":68C0
               Style           =   2  'Dropdown List
               TabIndex        =   377
               Top             =   2580
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   32
               ItemData        =   "frmParClinic.frx":68C2
               Left            =   705
               List            =   "frmParClinic.frx":68C4
               Style           =   2  'Dropdown List
               TabIndex        =   375
               Top             =   2250
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   22
               ItemData        =   "frmParClinic.frx":68C6
               Left            =   705
               List            =   "frmParClinic.frx":68C8
               Style           =   2  'Dropdown List
               TabIndex        =   363
               Top             =   270
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   23
               ItemData        =   "frmParClinic.frx":68CA
               Left            =   705
               List            =   "frmParClinic.frx":68CC
               Style           =   2  'Dropdown List
               TabIndex        =   365
               Top             =   600
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   24
               ItemData        =   "frmParClinic.frx":68CE
               Left            =   705
               List            =   "frmParClinic.frx":68D0
               Style           =   2  'Dropdown List
               TabIndex        =   367
               Top             =   930
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   25
               ItemData        =   "frmParClinic.frx":68D2
               Left            =   705
               List            =   "frmParClinic.frx":68D4
               Style           =   2  'Dropdown List
               TabIndex        =   369
               Top             =   1260
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   30
               ItemData        =   "frmParClinic.frx":68D6
               Left            =   705
               List            =   "frmParClinic.frx":68D8
               Style           =   2  'Dropdown List
               TabIndex        =   371
               Top             =   1590
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   31
               ItemData        =   "frmParClinic.frx":68DA
               Left            =   705
               List            =   "frmParClinic.frx":68DC
               Style           =   2  'Dropdown List
               TabIndex        =   373
               Top             =   1920
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   34
               ItemData        =   "frmParClinic.frx":68DE
               Left            =   705
               List            =   "frmParClinic.frx":68E0
               Style           =   2  'Dropdown List
               TabIndex        =   379
               Top             =   2910
               Width           =   2100
            End
            Begin VB.CheckBox chk 
               Caption         =   "����40�̶���С������ʾ"
               Height          =   180
               Index           =   119
               Left            =   135
               TabIndex        =   384
               Top             =   4170
               Width           =   2580
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   35
               ItemData        =   "frmParClinic.frx":68E2
               Left            =   705
               List            =   "frmParClinic.frx":68E4
               Style           =   2  'Dropdown List
               TabIndex        =   381
               Top             =   3240
               Width           =   2100
            End
            Begin VB.CheckBox chk 
               Caption         =   "�������(����Ϊ����42��)"
               Height          =   180
               Index           =   118
               Left            =   135
               TabIndex        =   383
               Top             =   3900
               Width           =   2655
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "˵����ʱ��֮����          ����"
               Height          =   180
               Index           =   26
               Left            =   135
               TabIndex        =   385
               Top             =   4455
               Width           =   2700
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Index           =   21
               Left            =   315
               TabIndex        =   376
               Top             =   2640
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Index           =   20
               Left            =   315
               TabIndex        =   374
               Top             =   2310
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ"
               Height          =   180
               Index           =   44
               Left            =   315
               TabIndex        =   362
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���"
               Height          =   180
               Index           =   45
               Left            =   315
               TabIndex        =   364
               Top             =   645
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ת��"
               Height          =   180
               Index           =   46
               Left            =   315
               TabIndex        =   366
               Top             =   990
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Index           =   48
               Left            =   315
               TabIndex        =   368
               Top             =   1320
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Index           =   49
               Left            =   315
               TabIndex        =   370
               Top             =   1650
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ"
               Height          =   180
               Index           =   50
               Left            =   315
               TabIndex        =   372
               Top             =   1980
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Index           =   19
               Left            =   315
               TabIndex        =   378
               Top             =   2970
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ת����"
               Height          =   180
               Index           =   18
               Left            =   135
               TabIndex        =   380
               Top             =   3300
               Width           =   540
            End
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   4110
            Index           =   1
            Left            =   135
            TabIndex        =   421
            Top             =   510
            Width           =   5535
            Begin VB.PictureBox picControl 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1845
               Index           =   0
               Left            =   3060
               ScaleHeight     =   1845
               ScaleWidth      =   2295
               TabIndex        =   292
               TabStop         =   0   'False
               Top             =   1215
               Visible         =   0   'False
               Width           =   2295
               Begin VB.PictureBox PicColorCollect 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1350
                  Index           =   0
                  Left            =   60
                  Picture         =   "frmParClinic.frx":68E6
                  ScaleHeight     =   1350
                  ScaleWidth      =   2160
                  TabIndex        =   293
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   2160
                  Begin VB.Shape shpBorder 
                     BorderColor     =   &H00C56A31&
                     FillColor       =   &H00FF8080&
                     Height          =   270
                     Index           =   0
                     Left            =   1890
                     Top             =   1080
                     Visible         =   0   'False
                     Width           =   270
                  End
                  Begin VB.Shape shpValue 
                     BorderColor     =   &H00C56A31&
                     FillColor       =   &H00FF8080&
                     Height          =   270
                     Index           =   0
                     Left            =   0
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   270
                  End
               End
               Begin VB.PictureBox picColor 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H80000008&
                  Height          =   200
                  Index           =   0
                  Left            =   90
                  ScaleHeight     =   165
                  ScaleWidth      =   165
                  TabIndex        =   294
                  TabStop         =   0   'False
                  Top             =   1575
                  Width           =   200
               End
               Begin VB.CommandButton cmdUnVisible 
                  Height          =   315
                  Index           =   0
                  Left            =   1815
                  Picture         =   "frmParClinic.frx":705C
                  Style           =   1  'Graphical
                  TabIndex        =   296
                  TabStop         =   0   'False
                  Top             =   1500
                  Width           =   450
               End
               Begin VB.Label lblColor 
                  Caption         =   "&HFFFFFF"
                  Height          =   195
                  Index           =   0
                  Left            =   405
                  TabIndex        =   295
                  Top             =   1575
                  UseMnemonic     =   0   'False
                  Width           =   1365
               End
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   3
               ItemData        =   "frmParClinic.frx":75E6
               Left            =   1860
               List            =   "frmParClinic.frx":75E8
               Style           =   2  'Dropdown List
               TabIndex        =   276
               Top             =   60
               Width           =   2565
            End
            Begin VB.CheckBox chk 
               Caption         =   "Ԥ������ӡʱǩ������ʾǩ��ͼƬ"
               Height          =   180
               Index           =   117
               Left            =   360
               TabIndex        =   288
               Top             =   2865
               Width           =   3645
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����ļ�ҳ�밴�ļ�˳����"
               Height          =   180
               Index           =   116
               Left            =   360
               TabIndex        =   291
               Top             =   3750
               Width           =   3135
            End
            Begin VB.CheckBox chk 
               Caption         =   "סԺ����ͬһʱ����Ҫ��¼��ݻ����ļ�"
               Height          =   180
               Index           =   115
               Left            =   360
               TabIndex        =   286
               Top             =   2280
               Width           =   3645
            End
            Begin VB.CheckBox chk 
               Caption         =   "ֻ�ڵ�ǰҳ����ʾ��ҳ���ݣ�������ҳ����ʾ��"
               Height          =   180
               Index           =   114
               Left            =   360
               TabIndex        =   290
               Top             =   3450
               Width           =   4215
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   14
               ItemData        =   "frmParClinic.frx":75EA
               Left            =   1860
               List            =   "frmParClinic.frx":75EC
               Style           =   2  'Dropdown List
               TabIndex        =   280
               Top             =   885
               Width           =   2565
            End
            Begin VB.TextBox txtUD 
               Alignment       =   2  'Center
               Height          =   270
               IMEMode         =   3  'DISABLE
               Index           =   21
               Left            =   1860
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   284
               Text            =   "1"
               Top             =   1920
               Width           =   375
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   13
               ItemData        =   "frmParClinic.frx":75EE
               Left            =   1860
               List            =   "frmParClinic.frx":75F0
               Style           =   2  'Dropdown List
               TabIndex        =   278
               Top             =   480
               Width           =   2565
            End
            Begin VB.CheckBox chk 
               Caption         =   "Ԥ������ӡʱͬһҳ��ͬ������ʾһ��"
               Height          =   180
               Index           =   113
               Left            =   360
               TabIndex        =   289
               Top             =   3165
               Width           =   3540
            End
            Begin VB.PictureBox picLineColor 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   1860
               ScaleHeight     =   180
               ScaleWidth      =   2535
               TabIndex        =   282
               TabStop         =   0   'False
               Top             =   1575
               Width           =   2565
            End
            Begin MSComCtl2.UpDown ud 
               Height          =   270
               Index           =   21
               Left            =   2235
               TabIndex        =   285
               TabStop         =   0   'False
               Top             =   1920
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   476
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txtUD(21)"
               BuddyDispid     =   196669
               BuddyIndex      =   21
               OrigLeft        =   2160
               OrigTop         =   1830
               OrigRight       =   2415
               OrigBottom      =   2100
               Max             =   30
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ͬ�����ļ�֮����������ͬ��"
               Height          =   180
               Index           =   167
               Left            =   645
               TabIndex        =   287
               Top             =   2580
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "����δ��¼�������ݵ��»���"
               Height          =   180
               Index           =   188
               Left            =   1890
               TabIndex        =   649
               Top             =   1260
               Width           =   2655
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ǩģʽ"
               Height          =   180
               Index           =   17
               Left            =   1080
               TabIndex        =   275
               Top             =   120
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "С��ȱʡ��ʶ"
               Height          =   180
               Index           =   16
               Left            =   720
               TabIndex        =   279
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����¼�볬����ǰ        ��Ļ����¼����"
               Height          =   180
               Index           =   15
               Left            =   360
               TabIndex        =   283
               Top             =   1950
               Width           =   3600
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ʿ��ǩ������ʾģʽ"
               Height          =   180
               Index           =   14
               Left            =   15
               TabIndex        =   277
               Top             =   540
               Width           =   1800
            End
            Begin VB.Label lblLineColor 
               AutoSize        =   -1  'True
               Caption         =   "С���ʶ��ɫ"
               Height          =   180
               Index           =   0
               Left            =   720
               TabIndex        =   281
               Top             =   1590
               Width           =   1080
            End
         End
         Begin VB.Frame fraPoliceLine 
            Height          =   1230
            Left            =   -69285
            TabIndex        =   414
            Top             =   1350
            Width           =   3735
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   5
               ItemData        =   "frmParClinic.frx":75F2
               Left            =   1590
               List            =   "frmParClinic.frx":75FC
               Style           =   2  'Dropdown List
               TabIndex        =   418
               Top             =   765
               Width           =   1905
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   4
               ItemData        =   "frmParClinic.frx":760C
               Left            =   1575
               List            =   "frmParClinic.frx":7616
               Style           =   2  'Dropdown List
               TabIndex        =   416
               Top             =   345
               Width           =   1905
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ͼ����ʾ���䡢�쳣��"
               Height          =   180
               Index           =   112
               Left            =   0
               TabIndex        =   413
               Top             =   0
               Width           =   2550
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "��������ʾΪ"
               Height          =   180
               Index           =   13
               Left            =   465
               TabIndex        =   415
               Top             =   390
               Width           =   1080
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "�쳣����ʾΪ"
               Height          =   180
               Index           =   12
               Left            =   465
               TabIndex        =   417
               Top             =   825
               Width           =   1080
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "��¶�ߵ���ʾ�����(����Ϊ�Ҳ�)"
            Height          =   180
            Index           =   111
            Left            =   -69285
            TabIndex        =   412
            Top             =   1065
            Width           =   3330
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ͼģʽΪ����ʽ(����Ϊ����ʽ)"
            Height          =   180
            Index           =   110
            Left            =   -69285
            TabIndex        =   411
            Top             =   795
            Width           =   3330
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ͼ����ʾ����ʱ��"
            Height          =   180
            Index           =   109
            Left            =   -69285
            TabIndex        =   410
            Top             =   525
            Width           =   2490
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   6
            ItemData        =   "frmParClinic.frx":7626
            Left            =   -67185
            List            =   "frmParClinic.frx":7630
            Style           =   2  'Dropdown List
            TabIndex        =   420
            Top             =   2685
            Width           =   1650
         End
         Begin VB.Frame fra2 
            Caption         =   "������ʩ��־"
            Height          =   1230
            Left            =   -74475
            TabIndex        =   405
            Top             =   3405
            Width           =   3735
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   9
               ItemData        =   "frmParClinic.frx":764C
               Left            =   1110
               List            =   "frmParClinic.frx":7659
               Style           =   2  'Dropdown List
               TabIndex        =   407
               Top             =   330
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   10
               ItemData        =   "frmParClinic.frx":7687
               Left            =   1110
               List            =   "frmParClinic.frx":7691
               Style           =   2  'Dropdown List
               TabIndex        =   409
               Top             =   765
               Width           =   2100
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��־����"
               Height          =   180
               Index           =   10
               Left            =   360
               TabIndex        =   406
               Top             =   390
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��־λ��"
               Height          =   180
               Index           =   9
               Left            =   360
               TabIndex        =   408
               Top             =   825
               Width           =   720
            End
         End
         Begin VB.Frame fra1 
            Caption         =   "�������߱�־(˳��)"
            Height          =   1230
            Index           =   0
            Left            =   -74475
            TabIndex        =   395
            Top             =   525
            Width           =   3735
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   12
               ItemData        =   "frmParClinic.frx":76AD
               Left            =   1110
               List            =   "frmParClinic.frx":76BA
               Style           =   2  'Dropdown List
               TabIndex        =   399
               Top             =   765
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   11
               ItemData        =   "frmParClinic.frx":76E4
               Left            =   1110
               List            =   "frmParClinic.frx":76F1
               Style           =   2  'Dropdown List
               TabIndex        =   397
               Top             =   330
               Width           =   2100
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��¶�½�"
               Height          =   180
               Index           =   8
               Left            =   360
               TabIndex        =   398
               Top             =   825
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               Height          =   180
               Index           =   5
               Left            =   360
               TabIndex        =   396
               Top             =   390
               Width           =   720
            End
         End
         Begin VB.Frame fra1 
            Caption         =   "�������߱�־(�쳣��)"
            Height          =   1230
            Index           =   1
            Left            =   -74475
            TabIndex        =   400
            Top             =   1965
            Width           =   3735
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   7
               ItemData        =   "frmParClinic.frx":771B
               Left            =   1110
               List            =   "frmParClinic.frx":7728
               Style           =   2  'Dropdown List
               TabIndex        =   402
               Top             =   330
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   8
               ItemData        =   "frmParClinic.frx":7752
               Left            =   1110
               List            =   "frmParClinic.frx":775F
               Style           =   2  'Dropdown List
               TabIndex        =   404
               Top             =   765
               Width           =   2100
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               Height          =   180
               Index           =   6
               Left            =   360
               TabIndex        =   401
               Top             =   390
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��¶�½�"
               Height          =   180
               Index           =   7
               Left            =   360
               TabIndex        =   403
               Top             =   825
               Width           =   720
            End
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   27
            Left            =   -72870
            TabIndex        =   299
            TabStop         =   0   'False
            Top             =   525
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(27)"
            BuddyDispid     =   196669
            BuddyIndex      =   27
            OrigLeft        =   2190
            OrigTop         =   870
            OrigRight       =   2430
            OrigBottom      =   1170
            Max             =   4
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   26
            Left            =   -72045
            TabIndex        =   308
            TabStop         =   0   'False
            Top             =   840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(26)"
            BuddyDispid     =   196669
            BuddyIndex      =   26
            OrigLeft        =   2190
            OrigTop         =   870
            OrigRight       =   2430
            OrigBottom      =   1170
            Max             =   30
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   24
            Left            =   -70320
            TabIndex        =   305
            TabStop         =   0   'False
            Top             =   525
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   8
            BuddyControl    =   "txtUD(24)"
            BuddyDispid     =   196669
            BuddyIndex      =   24
            OrigLeft        =   4580
            OrigTop         =   885
            OrigRight       =   4835
            OrigBottom      =   1170
            Max             =   23
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   23
            Left            =   -71670
            TabIndex        =   302
            TabStop         =   0   'False
            Top             =   525
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   18
            BuddyControl    =   "txtUD(23)"
            BuddyDispid     =   196669
            BuddyIndex      =   23
            OrigLeft        =   3230
            OrigTop         =   885
            OrigRight       =   3485
            OrigBottom      =   1170
            Max             =   23
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   22
            Left            =   -72045
            TabIndex        =   311
            TabStop         =   0   'False
            Top             =   1155
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtUD(22)"
            BuddyDispid     =   196669
            BuddyIndex      =   22
            OrigLeft        =   2190
            OrigTop         =   870
            OrigRight       =   2430
            OrigBottom      =   1170
            Max             =   50
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.PictureBox picControl 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1845
            Index           =   1
            Left            =   -70665
            ScaleHeight     =   1845
            ScaleWidth      =   2295
            TabIndex        =   390
            TabStop         =   0   'False
            Top             =   915
            Visible         =   0   'False
            Width           =   2295
            Begin VB.CommandButton cmdUnVisible 
               Height          =   315
               Index           =   1
               Left            =   1815
               Picture         =   "frmParClinic.frx":7789
               Style           =   1  'Graphical
               TabIndex        =   394
               TabStop         =   0   'False
               Top             =   1515
               Width           =   450
            End
            Begin VB.PictureBox PicColorCollect 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1350
               Index           =   1
               Left            =   60
               Picture         =   "frmParClinic.frx":7D13
               ScaleHeight     =   1350
               ScaleWidth      =   2160
               TabIndex        =   391
               TabStop         =   0   'False
               Top             =   90
               Width           =   2160
               Begin VB.Shape shpValue 
                  BorderColor     =   &H00C56A31&
                  FillColor       =   &H00FF8080&
                  Height          =   270
                  Index           =   1
                  Left            =   0
                  Top             =   -15
                  Visible         =   0   'False
                  Width           =   270
               End
               Begin VB.Shape shpBorder 
                  BorderColor     =   &H00C56A31&
                  FillColor       =   &H00FF8080&
                  Height          =   270
                  Index           =   1
                  Left            =   1890
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   270
               End
            End
            Begin VB.PictureBox picColor 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   200
               Index           =   1
               Left            =   90
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   392
               TabStop         =   0   'False
               Top             =   1575
               Width           =   200
            End
            Begin VB.Label lblColor 
               Caption         =   "&HFFFFFF"
               Height          =   195
               Index           =   1
               Left            =   405
               TabIndex        =   393
               Top             =   1575
               UseMnemonic     =   0   'False
               Width           =   1365
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "���µ��Ե�����ʾ(����Ϊ˫��)"
            Height          =   180
            Index           =   127
            Left            =   -68325
            TabIndex        =   357
            Top             =   585
            Width           =   2895
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�೦�����ʾ��ʽ��"
            Height          =   180
            Index           =   41
            Left            =   -74700
            TabIndex        =   353
            Top             =   6840
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�������ע    ��,��ɫ"
            Height          =   180
            Index           =   40
            Left            =   -74700
            TabIndex        =   325
            Top             =   3045
            Width           =   1890
         End
         Begin VB.Line Line1 
            X1              =   -73770
            X2              =   -73485
            Y1              =   3255
            Y2              =   3255
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ",����ȱʡ��ʽ"
            Height          =   180
            Index           =   27
            Left            =   -72390
            TabIndex        =   328
            Top             =   3045
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ��־������ʾ��"
            Height          =   180
            Index           =   39
            Left            =   -74700
            TabIndex        =   349
            Top             =   6585
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������¼�뷽ʽ��"
            Height          =   180
            Index           =   38
            Left            =   -74700
            TabIndex        =   345
            Top             =   6345
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���������䷽ʽ��"
            Height          =   180
            Index           =   37
            Left            =   -74700
            TabIndex        =   340
            Top             =   6090
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���¸��Ժϸ���ʾ����    ,��ɫ"
            Height          =   180
            Index           =   36
            Left            =   -74700
            TabIndex        =   322
            Top             =   2760
            Width           =   2610
         End
         Begin VB.Line Line2 
            X1              =   -72885
            X2              =   -72600
            Y1              =   2955
            Y2              =   2955
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   ",��ɫ"
            Height          =   180
            Index           =   35
            Left            =   -71595
            TabIndex        =   314
            Top             =   1515
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "δ��˵����ʾλ��"
            Height          =   180
            Index           =   34
            Left            =   -74700
            TabIndex        =   312
            Top             =   1530
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���²�����ʾ��ʽ"
            Height          =   180
            Index           =   33
            Left            =   -74700
            TabIndex        =   316
            Top             =   1860
            Width           =   1440
         End
         Begin VB.Label lblBreathe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϊ���ʱ,���������ݵ���ʾ��ʽ"
            Height          =   180
            Left            =   -74700
            TabIndex        =   320
            Top             =   2490
            Width           =   3150
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϊ���ʱ,�������ʱ����ʾ��ʽ"
            Height          =   180
            Index           =   32
            Left            =   -74700
            TabIndex        =   318
            Top             =   2175
            Width           =   3150
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���¿�ʼ��¼ʱ��"
            Height          =   180
            Index           =   31
            Left            =   -74700
            TabIndex        =   297
            Top             =   585
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���±����ʱ,������ݹ̶�        ��"
            Height          =   180
            Index           =   30
            Left            =   -74700
            TabIndex        =   306
            Top             =   885
            Width           =   3150
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����       ��"
            Height          =   180
            Index           =   24
            Left            =   -71010
            TabIndex        =   303
            Top             =   585
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ҹ���       ����"
            Height          =   180
            Index           =   23
            Left            =   -72540
            TabIndex        =   300
            Top             =   585
            Width           =   1530
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���±����ʱ,���߹̶����        ��"
            Height          =   180
            Index           =   22
            Left            =   -74700
            TabIndex        =   309
            Top             =   1185
            Width           =   3150
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����ͼ0�����һ�����ߵ�"
            Height          =   180
            Index           =   11
            Left            =   -69285
            TabIndex        =   419
            Top             =   2745
            Width           =   2070
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7335
      Index           =   5
      Left            =   2400
      ScaleHeight     =   7305
      ScaleWidth      =   10065
      TabIndex        =   118
      Top             =   0
      Width           =   10095
      Begin VB.CheckBox chk 
         Caption         =   "����ҩ�����հ�������������"
         Height          =   315
         Index           =   163
         Left            =   7335
         TabIndex        =   548
         Top             =   90
         Width           =   2700
      End
      Begin TabDlg.SSTab SST 
         Height          =   7095
         Left            =   120
         TabIndex        =   119
         Top             =   120
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   12515
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "����"
         TabPicture(0)   =   "frmParClinic.frx":8489
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl(28)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblDept(28)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "vsUseDept(28)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "vsfDrugStore(28)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cbo(28)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdDel(28)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdAdd(28)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "סԺ"
         TabPicture(1)   =   "frmParClinic.frx":84A5
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cbo(29)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmdDel(29)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdAdd(29)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "vsUseDept(29)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "vsfDrugStore(29)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lblDept(29)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lblDept(0)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "lbl(29)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).ControlCount=   8
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   29
            Left            =   -74280
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   452
            Width           =   1455
         End
         Begin VB.CommandButton cmdDel 
            Height          =   315
            Index           =   29
            Left            =   -72225
            Picture         =   "frmParClinic.frx":84C1
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   445
            Width           =   450
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   315
            Index           =   29
            Left            =   -72720
            Picture         =   "frmParClinic.frx":ED13
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   445
            Width           =   450
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   315
            Index           =   28
            Left            =   2280
            Picture         =   "frmParClinic.frx":15565
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   443
            Width           =   450
         End
         Begin VB.CommandButton cmdDel 
            Height          =   315
            Index           =   28
            Left            =   2775
            Picture         =   "frmParClinic.frx":1BDB7
            Style           =   1  'Graphical
            TabIndex        =   122
            Top             =   443
            Width           =   450
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   28
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   121
            Top             =   450
            Width           =   1455
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
            Height          =   5475
            Index           =   28
            Left            =   4440
            TabIndex        =   124
            Top             =   1200
            Width           =   5055
            _cx             =   8916
            _cy             =   9657
            Appearance      =   1
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
            BackColorBkg    =   14737632
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":22609
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
         Begin VSFlex8Ctl.VSFlexGrid vsUseDept 
            Height          =   5445
            Index           =   28
            Left            =   240
            TabIndex        =   125
            Top             =   1200
            Width           =   3885
            _cx             =   6853
            _cy             =   9604
            Appearance      =   1
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
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   245
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":226B6
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            Editable        =   2
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
         Begin VSFlex8Ctl.VSFlexGrid vsUseDept 
            Height          =   5445
            Index           =   29
            Left            =   -74760
            TabIndex        =   132
            Top             =   1200
            Width           =   3885
            _cx             =   6853
            _cy             =   9604
            Appearance      =   1
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   245
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":22862
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            Editable        =   2
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
         Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
            Height          =   5415
            Index           =   29
            Left            =   -70560
            TabIndex        =   134
            Top             =   1200
            Width           =   5055
            _cx             =   8916
            _cy             =   9551
            Appearance      =   1
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
            BackColorBkg    =   14737632
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":22A0E
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
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩҩ��"
            Height          =   180
            Index           =   29
            Left            =   -70560
            TabIndex        =   135
            Top             =   937
            Width           =   720
         End
         Begin VB.Label lblDept 
            Caption         =   "Ӧ�ÿ���"
            Height          =   255
            Index           =   0
            Left            =   -74760
            TabIndex        =   133
            Top             =   900
            Width           =   855
         End
         Begin VB.Label lbl 
            Caption         =   "����"
            Height          =   255
            Index           =   29
            Left            =   -74760
            TabIndex        =   131
            Top             =   490
            Width           =   375
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩҩ��"
            Height          =   180
            Index           =   28
            Left            =   4440
            TabIndex        =   127
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label8 
            Caption         =   "Ӧ�ÿ���"
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   900
            Width           =   855
         End
         Begin VB.Label lbl 
            Caption         =   "����"
            Height          =   255
            Index           =   28
            Left            =   240
            TabIndex        =   120
            Top             =   495
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8175
      Index           =   7
      Left            =   2400
      ScaleHeight     =   8145
      ScaleWidth      =   10065
      TabIndex        =   149
      Top             =   0
      Width           =   10095
      Begin VB.CheckBox chk 
         Caption         =   "�������¼�����Բ���д�����"
         Height          =   255
         Index           =   187
         Left            =   240
         TabIndex        =   648
         Top             =   4800
         Width           =   4455
      End
      Begin VB.CheckBox chk 
         Caption         =   "���¼��ʱ�����Զ���ȡ"
         Height          =   255
         Index           =   170
         Left            =   240
         TabIndex        =   630
         Top             =   4560
         Width           =   4455
      End
      Begin VB.CheckBox chk 
         Caption         =   "��ICD-10¼��ʱ���������ֻ����¼��M��ͷ��������̬ѧ����"
         Height          =   255
         Index           =   164
         Left            =   240
         TabIndex        =   549
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   542
         Top             =   1920
         Width           =   4215
         Begin VB.OptionButton optICD���� 
            Caption         =   "������д"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   545
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optICD���� 
            Caption         =   "��ʾ�Ƿ���д"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   544
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optICD���� 
            Caption         =   "�����"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   543
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "·������ԭ����ֵ����ѡȡ������������¼��"
         Height          =   255
         Index           =   155
         Left            =   2520
         TabIndex        =   521
         Top             =   4250
         Width           =   4215
      End
      Begin VB.CheckBox chk 
         Caption         =   "��������������ɵ���"
         Height          =   255
         Index           =   154
         Left            =   240
         TabIndex        =   520
         Top             =   4250
         Width           =   3375
      End
      Begin VB.CheckBox chk 
         Caption         =   "���֤������ʾ"
         Height          =   255
         Index           =   136
         Left            =   240
         TabIndex        =   484
         Top             =   615
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   168
         Top             =   1320
         Width           =   4215
         Begin VB.OptionButton opt������� 
            Caption         =   "�����"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   171
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton opt������� 
            Caption         =   "��ʾ�Ƿ���д"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   170
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton opt������� 
            Caption         =   "������д"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   169
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   164
         Top             =   960
         Width           =   3615
         Begin VB.OptionButton opt�����ж� 
            Caption         =   "������д"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   167
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton opt�����ж� 
            Caption         =   "��ʾ�Ƿ���д"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   166
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton opt�����ж� 
            Caption         =   "�����"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   165
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   160
         Top             =   2200
         Width           =   4215
         Begin VB.OptionButton opt���� 
            Caption         =   "������д"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   163
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "��ʾ�Ƿ���д"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   162
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "�����"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   161
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   159
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "ʹ����������ʱ��"
         Height          =   255
         Index           =   69
         Left            =   2280
         TabIndex        =   158
         Top             =   615
         Width           =   2055
      End
      Begin VB.CheckBox chk 
         Caption         =   $"frmParClinic.frx":22A97
         Height          =   255
         Index           =   70
         Left            =   240
         TabIndex        =   157
         Top             =   2475
         Width           =   4095
      End
      Begin VB.CheckBox chk 
         Caption         =   "ҽ���ͻ�ʿ�ֱ���д������ҳ"
         Height          =   255
         Index           =   71
         Left            =   240
         TabIndex        =   156
         Top             =   3480
         Width           =   4095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   8685
         TabIndex        =   152
         Top             =   5175
         Width           =   1100
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "�޸�(&M)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7485
         TabIndex        =   151
         Top             =   5175
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddMed 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   6285
         TabIndex        =   150
         Top             =   5175
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMecItem 
         Height          =   2415
         Left            =   120
         TabIndex        =   153
         Top             =   5640
         Width           =   9720
         _cx             =   17145
         _cy             =   4260
         Appearance      =   3
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
         BackColorSel    =   16574424
         ForeColorSel    =   -2147483642
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
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
      Begin VB.Label lblICD���� 
         AutoSize        =   -1  'True
         Caption         =   "��Ҫ��Ժ���ΪC00��D48ʱ,ICD���룺"
         Height          =   180
         Left            =   240
         TabIndex        =   546
         Top             =   1935
         Width           =   3240
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         Caption         =   "��Ҫ��Ժ���ΪC00��D48ʱ,������ϣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   177
         Top             =   1335
         Width           =   3330
      End
      Begin VB.Label lbl�����ж� 
         AutoSize        =   -1  'True
         Caption         =   "��Ҫ��Ժ��ϱ���ΪS��T��ʱ,�����ж���ϣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   176
         Top             =   975
         Width           =   3690
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ�����"
         Height          =   180
         Left            =   240
         TabIndex        =   175
         Top             =   2220
         Width           =   1260
      End
      Begin VB.Label lbl��ҳ��׼ 
         Caption         =   "������ҳ��׼"
         Height          =   255
         Left            =   240
         TabIndex        =   174
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label lbl��ҽ 
         Caption         =   $"frmParClinic.frx":22AC3
         Height          =   585
         Left            =   480
         TabIndex        =   173
         Top             =   2790
         Width           =   6015
      End
      Begin VB.Label lblSeparEdit 
         Caption         =   "�����¼�����Ŀ����Һ��Ӧ������ҩ��ٴ����֡�סԺ�ڼ�����Լ������Ժʱ͸��(Ѫ͸����͸)���ص�ֵ����Ϣ�����øò���ʱֻ���ɻ�ʿ��д"
         Height          =   360
         Left            =   480
         TabIndex        =   172
         Top             =   3800
         Width           =   6015
      End
      Begin VB.Label Label3 
         Caption         =   "סԺ��ҳ������Ŀ��"
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   5280
         Width           =   4095
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7725
      Index           =   0
      Left            =   2400
      ScaleHeight     =   7695
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CheckBox chk 
         Caption         =   "ҽ����ҩ��������"
         Height          =   240
         Index           =   138
         Left            =   240
         TabIndex        =   486
         Top             =   2445
         Width           =   1875
      End
      Begin VB.CommandButton cmdAdvice 
         Caption         =   "ҽ�����ݶ���(&F)"
         Height          =   405
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   1680
      End
      Begin VB.Frame fra��Ժ��� 
         Caption         =   "סԺ�´��������ҽ��ʱ����Ƿ���д��Ժ���"
         Height          =   1365
         Left            =   240
         TabIndex        =   22
         Top             =   6165
         Width           =   4320
         Begin VB.CommandButton cmdסԺ�����Ժ��� 
            Caption         =   "ȫ��"
            Height          =   300
            Index           =   1
            Left            =   3120
            TabIndex        =   25
            Top             =   720
            Width           =   900
         End
         Begin VB.CommandButton cmdסԺ�����Ժ��� 
            Caption         =   "ȫѡ"
            Height          =   300
            Index           =   0
            Left            =   3120
            TabIndex        =   24
            Top             =   360
            Width           =   900
         End
         Begin VB.ListBox lst 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   375
            Width           =   2940
         End
      End
      Begin VB.Frame fra����Ŀ�� 
         BorderStyle     =   0  'None
         Height          =   1935
         Index           =   2
         Left            =   240
         TabIndex        =   422
         Top             =   3600
         Width           =   4380
         Begin VB.Frame fra����Ŀ�� 
            Caption         =   "����"
            Height          =   680
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   330
            Width           =   4305
            Begin VB.OptionButton opt����Ŀ������ 
               Caption         =   "����"
               Height          =   180
               Index           =   2
               Left            =   3000
               TabIndex        =   17
               Top             =   300
               Value           =   -1  'True
               Width           =   680
            End
            Begin VB.OptionButton opt����Ŀ������ 
               Caption         =   "Ԥ��"
               Height          =   180
               Index           =   1
               Left            =   1920
               TabIndex        =   16
               Top             =   300
               Width           =   680
            End
            Begin VB.OptionButton opt����Ŀ������ 
               Caption         =   "�´�ʱȷ��"
               Height          =   180
               Index           =   0
               Left            =   255
               TabIndex        =   15
               Top             =   300
               Width           =   1275
            End
         End
         Begin VB.Frame fra����Ŀ�� 
            Caption         =   "סԺ"
            Height          =   680
            Index           =   1
            Left            =   0
            TabIndex        =   18
            Top             =   1200
            Width           =   4305
            Begin VB.OptionButton opt����Ŀ��סԺ 
               Caption         =   "�´�ʱȷ��"
               Height          =   180
               Index           =   0
               Left            =   255
               TabIndex        =   19
               Top             =   300
               Width           =   1275
            End
            Begin VB.OptionButton opt����Ŀ��סԺ 
               Caption         =   "Ԥ��"
               Height          =   180
               Index           =   1
               Left            =   1920
               TabIndex        =   20
               Top             =   300
               Width           =   680
            End
            Begin VB.OptionButton opt����Ŀ��סԺ 
               Caption         =   "����"
               Height          =   180
               Index           =   2
               Left            =   3000
               TabIndex        =   21
               Top             =   300
               Value           =   -1  'True
               Width           =   680
            End
         End
         Begin VB.Label lbl����Ŀ�� 
            Caption         =   "����ҩ��ȱʡ��ҩĿ��"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   60
            Width           =   1935
         End
      End
      Begin VB.Frame fraҽ���´� 
         Caption         =   "סԺҽ���´�"
         Height          =   4170
         Left            =   4920
         TabIndex        =   423
         Top             =   3360
         Width           =   4770
         Begin VB.CheckBox chk 
            Caption         =   "Ƥ��������ҩ"
            Height          =   195
            Index           =   178
            Left            =   2775
            TabIndex        =   625
            Top             =   3120
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´ﶾ��͵�һ�ྫ��ҩƷʱ����ǼǴ�����"
            Height          =   195
            Index           =   174
            Left            =   210
            TabIndex        =   621
            Top             =   1150
            Width           =   3960
         End
         Begin VB.CheckBox chk 
            Caption         =   "����Ƥ�Խ������ҽ������"
            Height          =   225
            Index           =   166
            Left            =   210
            TabIndex        =   560
            Top             =   3120
            Width           =   2585
         End
         Begin VB.CheckBox chk 
            Caption         =   "סԺҽ���´�ʱ����ѡ������ʾҩƷ���"
            Height          =   240
            Index           =   51
            Left            =   210
            TabIndex        =   424
            Top             =   2550
            Width           =   4125
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҽ��ȱʡΪ������Ч"
            Height          =   240
            Index           =   24
            Left            =   210
            TabIndex        =   425
            Top             =   300
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷ����ҽ��������´�"
            Height          =   240
            Index           =   4
            Left            =   210
            TabIndex        =   426
            Top             =   570
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´�ҩƷ����ʱ����ָ����ҩ����"
            Height          =   195
            Index           =   15
            Left            =   210
            TabIndex        =   427
            Top             =   1440
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "������ִ��Ƶ��ȱʡΪһ����"
            Height          =   195
            Index           =   13
            Left            =   210
            TabIndex        =   428
            Top             =   870
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´��Ժҽ��ʱ����Ժ��ϵ���д"
            Height          =   195
            Index           =   16
            Left            =   210
            TabIndex        =   429
            Top             =   1725
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´�����ʱ�����뵥��"
            Height          =   195
            Index           =   14
            Left            =   2280
            TabIndex        =   430
            Top             =   2850
            Width           =   2205
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ִ����ɺ�������´�����ҽ��"
            Height          =   195
            Index           =   17
            Left            =   210
            TabIndex        =   431
            Top             =   2010
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҽ��ʱ�Զ�����"
            Height          =   255
            Index           =   20
            Left            =   210
            TabIndex        =   432
            Top             =   3780
            Width           =   1935
         End
         Begin VB.CommandButton cmdAdviceSortSet 
            Caption         =   "�����������(&S)"
            Height          =   350
            Left            =   2280
            TabIndex        =   433
            Top             =   3720
            Width           =   1695
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������ס�����´�ҽ��"
            Height          =   195
            Index           =   18
            Left            =   210
            TabIndex        =   434
            Top             =   2280
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chk 
            Caption         =   "�Զ�����Ƥ��ҽ��"
            Height          =   225
            Index           =   19
            Left            =   210
            TabIndex        =   435
            Top             =   2850
            Width           =   1860
         End
         Begin VB.OptionButton optδƤ������ҽ�� 
            Caption         =   "����ҩƷ"
            Height          =   255
            Index           =   0
            Left            =   945
            TabIndex        =   436
            Top             =   3405
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optδƤ������ҽ�� 
            Caption         =   "�������Ľ��յ�ҩƷ"
            Height          =   255
            Index           =   1
            Left            =   2025
            TabIndex        =   437
            Top             =   3405
            Width           =   1920
         End
         Begin VB.Label lblSTCheck 
            Caption         =   "����"
            Height          =   255
            Left            =   480
            TabIndex        =   438
            Top             =   3420
            Width           =   375
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "����ҽ���´�"
         Height          =   3030
         Left            =   4920
         TabIndex        =   26
         Top             =   240
         Width           =   4770
         Begin VB.CheckBox chk 
            Caption         =   "������ҽ������¼����ҽ���"
            Height          =   180
            Index           =   186
            Left            =   240
            TabIndex        =   647
            Top             =   2715
            Width           =   3885
         End
         Begin VB.CheckBox chk 
            Caption         =   "Ƥ��������ҩ"
            Height          =   195
            Index           =   177
            Left            =   2760
            TabIndex        =   624
            Top             =   1335
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "����Ƥ�Խ������ҽ������"
            Height          =   195
            Index           =   165
            Left            =   240
            TabIndex        =   559
            Top             =   1335
            Width           =   2535
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҽ���´�ʱ����ѡ������ʾҩƷ���"
            Height          =   240
            Index           =   48
            Left            =   240
            TabIndex        =   439
            Top             =   2460
            Width           =   3660
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   3
            Left            =   2370
            TabIndex        =   440
            TabStop         =   0   'False
            Top             =   1890
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "txtUD(3)"
            BuddyDispid     =   196669
            BuddyIndex      =   3
            OrigLeft        =   2400
            OrigTop         =   1080
            OrigRight       =   2655
            OrigBottom      =   1350
            Max             =   999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   0   'False
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   1970
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   441
            Text            =   "5"
            Top             =   1890
            Width           =   375
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   1350
            MaxLength       =   4
            TabIndex        =   28
            Text            =   "10"
            Top             =   240
            Width           =   400
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҩ�������Ϻ���ҩ"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   442
            Top             =   2205
            Width           =   2460
         End
         Begin VB.CheckBox chk 
            Caption         =   "�Զ�����Ƥ��ҽ��"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   443
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´�ҩƷҽ��ʱ����ָ����ҩ����"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   444
            Top             =   825
            Width           =   3360
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´�ҩƷҽ��ʱ����¼��ҩƷ����"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   445
            Top             =   555
            Width           =   3360
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´ﶾ��͵�һ�ྫ��ҩƷʱ����ǼǴ�����"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   446
            Top             =   1620
            Width           =   3960
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   11
            Left            =   1760
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   10
            BuddyControl    =   "txtUD(11)"
            BuddyDispid     =   196669
            BuddyIndex      =   11
            OrigLeft        =   2100
            OrigTop         =   420
            OrigRight       =   2355
            OrigBottom      =   690
            Max             =   9999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chk 
            Caption         =   "һ�Ŵ����������        ��ҩƷҽ��"
            Height          =   240
            Index           =   52
            Left            =   240
            TabIndex        =   447
            Top             =   1905
            Width           =   3780
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�¿�ҽ�����        �������Ե�ǰʱ����Ϊ��ʼʱ��"
            Height          =   180
            Index           =   25
            Left            =   240
            TabIndex        =   27
            Top             =   285
            Width           =   4320
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "���������Һ���Ч�����Ĳ���"
         Height          =   195
         Index           =   82
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   3045
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����Ǽ���Ч����"
         Height          =   240
         Index           =   11
         Left            =   240
         TabIndex        =   4
         Top             =   690
         Width           =   1740
      End
      Begin VB.CheckBox chk 
         Caption         =   "һ��������������Ŀ"
         Height          =   240
         Index           =   34
         Left            =   240
         TabIndex        =   10
         Top             =   1860
         Width           =   2820
      End
      Begin VB.TextBox txtUD 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   270
         Index           =   7
         Left            =   2010
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "1"
         Top             =   675
         Width           =   495
      End
      Begin VB.CheckBox chk 
         Caption         =   "�´�ҽ��ʱ��ʾ����"
         Height          =   240
         Index           =   66
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   26
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtUD 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   1710
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "30"
         Top             =   240
         Width           =   465
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   8
         Left            =   2205
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   30
         BuddyControl    =   "txtUD(8)"
         BuddyDispid     =   196669
         BuddyIndex      =   8
         OrigLeft        =   2100
         OrigTop         =   120
         OrigRight       =   2340
         OrigBottom      =   390
         Max             =   9999
         Min             =   10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   7
         Left            =   2520
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   675
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtUD(7)"
         BuddyDispid     =   196669
         BuddyIndex      =   7
         OrigLeft        =   2400
         OrigTop         =   1380
         OrigRight       =   2655
         OrigBottom      =   1680
         Max             =   365
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.Label lbl 
         Caption         =   "��ҩ�䷽ÿ��"
         Height          =   255
         Index           =   54
         Left            =   240
         TabIndex        =   7
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��¼ҽ��ʶ����         ����"
         Height          =   180
         Index           =   43
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   2610
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   3
      Left            =   2400
      ScaleHeight     =   7425
      ScaleWidth      =   10065
      TabIndex        =   112
      Top             =   0
      Width           =   10095
      Begin VB.Frame fraStopDept 
         Height          =   3215
         Left            =   5880
         TabIndex        =   561
         Top             =   4080
         Width           =   4095
         Begin VB.CheckBox chk 
            Caption         =   "ͣ��ʱ¼��ԭ��"
            Height          =   195
            Index           =   56
            Left            =   120
            TabIndex        =   562
            Top             =   0
            Width           =   1575
         End
         Begin VSFlex8Ctl.VSFlexGrid vsStopDept 
            Height          =   2640
            Left            =   120
            TabIndex        =   563
            Top             =   490
            Width           =   3900
            _cx             =   6879
            _cy             =   4657
            Appearance      =   2
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   9
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   500
            ColWidthMax     =   960
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":22B81
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
            Editable        =   2
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
         Begin VB.Label lblStopDept 
            AutoSize        =   -1  'True
            Caption         =   "���ÿɲ���дͣ��ԭ��Ŀ��ң����磺����ơ�"
            Height          =   180
            Left            =   120
            TabIndex        =   564
            Top             =   240
            Width           =   3780
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "����ʱ����𽫱���ִ�е���Ŀ��Ϊ��ִ��"
         Height          =   3915
         Left            =   5880
         TabIndex        =   494
         Top             =   120
         Width           =   4065
         Begin VB.CommandButton cmdAdd 
            Height          =   315
            Index           =   37
            Left            =   2205
            Picture         =   "frmParClinic.frx":22C37
            Style           =   1  'Graphical
            TabIndex        =   640
            Top             =   1620
            Width           =   450
         End
         Begin VB.CommandButton cmdDel 
            Height          =   315
            Index           =   37
            Left            =   2700
            Picture         =   "frmParClinic.frx":29489
            Style           =   1  'Graphical
            TabIndex        =   639
            Top             =   1620
            Width           =   450
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   37
            Left            =   645
            Style           =   2  'Dropdown List
            TabIndex        =   638
            Top             =   1620
            Width           =   1455
         End
         Begin VB.CommandButton cmd����ִ���Զ����ҽ����� 
            Caption         =   "ȫ��"
            Height          =   300
            Index           =   1
            Left            =   3195
            TabIndex        =   499
            Top             =   855
            Width           =   800
         End
         Begin VB.CommandButton cmd����ִ���Զ����ҽ����� 
            Caption         =   "ȫѡ"
            Height          =   300
            Index           =   0
            Left            =   3195
            TabIndex        =   498
            Top             =   495
            Width           =   800
         End
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   255
            Index           =   40
            Left            =   975
            TabIndex        =   497
            Top             =   210
            Width           =   735
         End
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   255
            Index           =   39
            Left            =   135
            TabIndex        =   496
            Top             =   210
            Width           =   735
         End
         Begin VB.ListBox lst 
            Columns         =   3
            Height          =   900
            Index           =   8
            ItemData        =   "frmParClinic.frx":2FCDB
            Left            =   120
            List            =   "frmParClinic.frx":2FCE2
            Style           =   1  'Checkbox
            TabIndex        =   495
            Top             =   510
            Width           =   2985
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUseDept 
            Height          =   1650
            Index           =   37
            Left            =   135
            TabIndex        =   641
            Top             =   2190
            Width           =   3795
            _cx             =   6694
            _cy             =   2910
            Appearance      =   1
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
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   245
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":2FCF0
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            Editable        =   2
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   42
            Left            =   165
            TabIndex        =   643
            Top             =   1665
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Ӧ�ÿ���"
            Height          =   180
            Index           =   47
            Left            =   150
            TabIndex        =   642
            Top             =   1965
            Width           =   720
         End
      End
      Begin VB.Frame fraסԺҽ������ӡ 
         Caption         =   "סԺҽ������ӡģʽ"
         Height          =   900
         Left            =   150
         TabIndex        =   77
         Top             =   6395
         Width           =   2700
         Begin VB.OptionButton optסԺҽ������ӡ 
            Caption         =   "�¿�ʱ��ӡ"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   79
            Top             =   615
            Width           =   1440
         End
         Begin VB.OptionButton optסԺҽ������ӡ 
            Caption         =   "У�Ժ��ӡ"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   78
            Top             =   270
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.Frame fra��Ѫ���뵥��ӡ 
         Caption         =   "��Ѫ���뵥��ӡģʽ"
         Height          =   900
         Left            =   3015
         TabIndex        =   80
         Top             =   6395
         Width           =   2700
         Begin VB.OptionButton opt��Ѫ���뵥��ӡ 
            Caption         =   "����ʱ��ӡ"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   81
            Top             =   270
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.OptionButton opt��Ѫ���뵥��ӡ 
            Caption         =   "�¿�ʱ��ӡ"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   82
            Top             =   585
            Width           =   1440
         End
      End
      Begin VB.Frame fra�������� 
         Caption         =   "סԺҽ����������� "
         Height          =   2745
         Left            =   120
         TabIndex        =   67
         Top             =   3615
         Width           =   5625
         Begin VB.CheckBox chk 
            Caption         =   "��������´�ҽ���ɻ���������Ҵ���"
            Height          =   180
            Index           =   184
            Left            =   1980
            TabIndex        =   645
            Top             =   2415
            Width           =   3540
         End
         Begin VB.CheckBox chk 
            Caption         =   "������Ҫ����ִ��"
            Height          =   240
            Index           =   168
            Left            =   210
            TabIndex        =   565
            Top             =   2400
            Width           =   1845
         End
         Begin VB.CheckBox chk 
            Caption         =   "���˳�Ժҽ�����ܳ���Ԥ��Ժ"
            Height          =   240
            Index           =   81
            Left            =   2520
            TabIndex        =   76
            Top             =   2070
            Width           =   2700
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´��Ժҽ�����ܳ�Ժ"
            Height          =   240
            Index           =   50
            Left            =   210
            TabIndex        =   75
            Top             =   2070
            Width           =   2340
         End
         Begin VB.CheckBox chk 
            Caption         =   $"frmParClinic.frx":2FE9C
            Height          =   195
            Index           =   33
            Left            =   210
            TabIndex        =   74
            Top             =   1785
            Width           =   2460
         End
         Begin VB.CheckBox chk 
            Caption         =   "��������У��"
            Height          =   195
            Index           =   28
            Left            =   210
            TabIndex        =   70
            Top             =   907
            Width           =   1380
         End
         Begin VB.CheckBox chk 
            Caption         =   "����������ͣ/����"
            Height          =   195
            Index           =   29
            Left            =   2520
            TabIndex        =   71
            Top             =   907
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "У��,ȷ��ֹͣ,����ҽ������д�ӡ"
            Height          =   180
            Index           =   26
            Left            =   210
            TabIndex        =   68
            Top             =   360
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����ҽ���´��ҽ�����к�������"
            Height          =   195
            Index           =   31
            Left            =   210
            TabIndex        =   73
            Top             =   1515
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "У�Ժ�ȷ��ֹͣʱʹ�õ���ǩ��"
            Height          =   195
            Index           =   27
            Left            =   210
            TabIndex        =   69
            Top             =   626
            Width           =   3165
         End
         Begin VB.CheckBox chk 
            Caption         =   "��дƤ�Խ��ʱ��֤���"
            Height          =   195
            Index           =   30
            Left            =   210
            TabIndex        =   72
            Top             =   1230
            Width           =   2445
         End
      End
      Begin VB.Frame fraסԺ���� 
         Caption         =   "סԺ����ѡ��"
         Height          =   3450
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Width           =   5610
         Begin VB.CheckBox chk 
            Caption         =   "����δ����ҽ��ʱ��ֹ����ת��ҽ��"
            Height          =   180
            Index           =   185
            Left            =   240
            TabIndex        =   646
            Top             =   3195
            Width           =   4755
         End
         Begin VB.CheckBox chk 
            Caption         =   "������ɺ�ر�ҽ������"
            Height          =   180
            Index           =   161
            Left            =   240
            TabIndex        =   533
            Top             =   2970
            Width           =   3285
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҩƷ�ֿ�����"
            Height          =   180
            Index           =   160
            Left            =   240
            TabIndex        =   532
            Top             =   2720
            Width           =   1845
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҽ������ʱһ����鷢��Ϊһ�ŵ���"
            Height          =   180
            Index           =   141
            Left            =   240
            TabIndex        =   493
            Top             =   2420
            Width           =   4305
         End
         Begin VB.Frame Frame13 
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   255
            Left            =   2085
            TabIndex        =   488
            Top             =   2140
            Width           =   2655
            Begin VB.OptionButton opt��ҩ���� 
               Caption         =   "��ҩִ�п���"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   489
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opt��ҩ���� 
               Caption         =   "���˲���"
               Height          =   180
               Index           =   1
               Left            =   1560
               TabIndex        =   490
               Top             =   0
               Value           =   -1  'True
               Width           =   1050
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "סԺҩ�����Ͳ�����ҩ��"
            Height          =   240
            Index           =   64
            Left            =   240
            TabIndex        =   60
            Top             =   300
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����ڷ�ҩ���ͽ���ʱ��"
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   65
            Top             =   1820
            Width           =   2325
         End
         Begin VB.CheckBox chk 
            Caption         =   "��δУ�Ի������ҽ����ֹ����ת�ơ���Ժ��תԺ������ҽ��"
            Height          =   180
            Index           =   22
            Left            =   240
            TabIndex        =   63
            Top             =   1200
            Width           =   5220
         End
         Begin VB.CheckBox chk 
            Caption         =   "����У�Լ��ɷ���ҽ��"
            Height          =   180
            Index           =   21
            Left            =   240
            TabIndex        =   61
            Top             =   600
            Width           =   2160
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʱ��ҽ�����˼����Ŀ�Ƿ�����"
            Height          =   195
            Index           =   38
            Left            =   240
            TabIndex        =   62
            Top             =   900
            Value           =   1  'Checked
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷ�����ĸ�ҩ;�����ʹ����Խ���ʱ��Ϊ׼����"
            Height          =   180
            Index           =   23
            Left            =   240
            TabIndex        =   64
            Top             =   1500
            Width           =   4305
         End
         Begin MSComCtl2.DTPicker dtp�ڷ�����ʱ�� 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-M-d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   2565
            TabIndex        =   66
            Top             =   1767
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   159711234
            CurrentDate     =   0.999988425925926
         End
         Begin VB.Label lbl��ҩ���� 
            Caption         =   "ҩƷҽ������ҩ����Ϊ"
            Height          =   180
            Left            =   240
            TabIndex        =   491
            Top             =   2140
            Width           =   1845
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7875
      Index           =   1
      Left            =   2400
      ScaleHeight     =   7845
      ScaleWidth      =   10185
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   10215
      Begin VB.Frame FrmBloodManager 
         Height          =   1725
         Left            =   120
         TabIndex        =   622
         Top             =   6075
         Width           =   3465
         Begin VB.CheckBox chk 
            Caption         =   "ѪҺ���ҽ��պ��������дִ�еǼ�"
            Enabled         =   0   'False
            Height          =   240
            Index           =   183
            Left            =   120
            TabIndex        =   482
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3210
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´���Ѫ����ʱȷ����Ѫ��Ϣ"
            Enabled         =   0   'False
            Height          =   240
            Index           =   181
            Left            =   120
            TabIndex        =   481
            Top             =   810
            Width           =   2850
         End
         Begin VB.CheckBox chk 
            Caption         =   "����Ѫ�����ϵͳ"
            Height          =   240
            Index           =   135
            Left            =   120
            TabIndex        =   478
            Top             =   0
            Width           =   1785
         End
         Begin VB.ComboBox cbo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   36
            ItemData        =   "frmParClinic.frx":2FEBA
            Left            =   2280
            List            =   "frmParClinic.frx":2FEBC
            Style           =   2  'Dropdown List
            TabIndex        =   483
            Top             =   1365
            Width           =   1065
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ѫҽ�����ͺ���Ѫ�Ʋ��ܽ��з�Ѫ"
            Enabled         =   0   'False
            Height          =   255
            Index           =   176
            Left            =   120
            TabIndex        =   479
            Top             =   270
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ѫ���벻��ʾѪҺ�����Ϣ"
            Enabled         =   0   'False
            Height          =   240
            Index           =   175
            Left            =   120
            TabIndex        =   480
            Top             =   540
            Width           =   2850
         End
         Begin VB.Label lblBloodManager 
            Caption         =   "��Ѫ�ɼ�Ĭ�ϼ�����������"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   623
            Top             =   1425
            Width           =   2235
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "סԺ"
         Height          =   225
         Index           =   159
         Left            =   8175
         TabIndex        =   530
         Top             =   5085
         Width           =   675
      End
      Begin VB.CheckBox chk 
         Caption         =   "����"
         Height          =   225
         Index           =   158
         Left            =   7230
         TabIndex        =   529
         Top             =   5085
         Width           =   675
      End
      Begin VB.CheckBox chk 
         Caption         =   "��ƻ����ɻ�����������дһ�����"
         Height          =   240
         Index           =   146
         Left            =   3765
         TabIndex        =   509
         Top             =   7500
         Width           =   3420
      End
      Begin VB.Frame Frame1 
         Caption         =   "���뵥���û���"
         Height          =   900
         Index           =   5
         Left            =   3765
         TabIndex        =   503
         Top             =   4080
         Width           =   6225
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   225
            Index           =   157
            Left            =   3495
            TabIndex        =   528
            Top             =   570
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ѫ"
            Height          =   225
            Index           =   156
            Left            =   2595
            TabIndex        =   527
            Top             =   570
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   225
            Index           =   152
            Left            =   1680
            TabIndex        =   526
            Top             =   570
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "���"
            Height          =   225
            Index           =   151
            Left            =   765
            TabIndex        =   522
            Top             =   570
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "���"
            Height          =   225
            Index           =   145
            Left            =   765
            TabIndex        =   508
            Top             =   285
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   225
            Index           =   144
            Left            =   1680
            TabIndex        =   507
            Top             =   285
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ѫ"
            Height          =   225
            Index           =   143
            Left            =   2595
            TabIndex        =   506
            Top             =   285
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   225
            Index           =   142
            Left            =   3495
            TabIndex        =   505
            Top             =   285
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   225
            Index           =   90
            Left            =   4380
            TabIndex        =   504
            Top             =   570
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "סԺ��"
            Height          =   180
            Left            =   165
            TabIndex        =   525
            Top             =   570
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "���"
            Height          =   180
            Left            =   165
            TabIndex        =   524
            Top             =   285
            Width           =   540
         End
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         ItemData        =   "frmParClinic.frx":2FEBE
         Left            =   1920
         List            =   "frmParClinic.frx":2FEC0
         Style           =   2  'Dropdown List
         TabIndex        =   147
         Top             =   5715
         Width           =   1065
      End
      Begin VB.TextBox txt 
         Height          =   600
         Index           =   1
         Left            =   3765
         MultiLine       =   -1  'True
         TabIndex        =   448
         Top             =   6675
         Width           =   6255
      End
      Begin VB.TextBox txt 
         Height          =   600
         Index           =   0
         Left            =   3765
         MultiLine       =   -1  'True
         TabIndex        =   449
         Top             =   5715
         Width           =   6255
      End
      Begin VB.CheckBox chk 
         Caption         =   $"frmParClinic.frx":2FEC2
         Height          =   255
         Index           =   79
         Left            =   120
         TabIndex        =   450
         Top             =   5415
         Width           =   3000
      End
      Begin VB.Frame fraCLKS 
         Height          =   3840
         Left            =   3760
         TabIndex        =   451
         Top             =   120
         Width           =   6250
         Begin VB.CheckBox chk 
            Caption         =   "ҽ������ʱ��������ԭ��"
            Height          =   240
            Index           =   86
            Left            =   105
            TabIndex        =   452
            Top             =   0
            Width           =   2280
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUnWriteDept 
            Height          =   3165
            Left            =   120
            TabIndex        =   453
            Top             =   525
            Width           =   6000
            _cx             =   10583
            _cy             =   5583
            Appearance      =   2
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":2FEE6
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
            Editable        =   2
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
         Begin VB.Label Label16 
            Caption         =   "���ÿɲ�¼�볬��ԭ��Ŀ��ң����磺����ơ�"
            Height          =   255
            Left            =   240
            TabIndex        =   454
            Top             =   300
            Width           =   4095
         End
      End
      Begin VB.Frame frmOPS 
         Height          =   1095
         Left            =   120
         TabIndex        =   455
         Top             =   3315
         Width           =   3465
         Begin VB.CheckBox chk 
            Caption         =   "����ҽʦ�ﵽ�����ȼ��������"
            Enabled         =   0   'False
            Height          =   240
            Index           =   83
            Left            =   120
            TabIndex        =   519
            Top             =   780
            Width           =   3105
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����ּ����"
            Enabled         =   0   'False
            Height          =   240
            Index           =   140
            Left            =   120
            TabIndex        =   492
            Top             =   525
            Width           =   1425
         End
         Begin VB.CheckBox chk 
            Caption         =   "��������ҽʦ��Ȩ����"
            Enabled         =   0   'False
            Height          =   240
            Index           =   49
            Left            =   120
            TabIndex        =   456
            Top             =   270
            Width           =   2220
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������ּ�����"
            Height          =   240
            Index           =   80
            Left            =   120
            TabIndex        =   457
            Top             =   0
            Width           =   1740
         End
      End
      Begin VB.Frame fraKSSStrict 
         Height          =   915
         Index           =   14
         Left            =   120
         TabIndex        =   458
         Top             =   2340
         Width           =   3465
         Begin VB.CheckBox chk 
            Caption         =   "��ҽ��С����п���ҩ�����"
            Enabled         =   0   'False
            Height          =   255
            Index           =   137
            Left            =   120
            TabIndex        =   485
            Top             =   585
            Width           =   2700
         End
         Begin VB.CheckBox chk 
            Caption         =   "���ÿ���ҩ��ּ�����"
            Height          =   255
            Index           =   75
            Left            =   120
            TabIndex        =   459
            Top             =   0
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҩ������ʹ���Ա�ҩ"
            Enabled         =   0   'False
            Height          =   255
            Index           =   76
            Left            =   120
            TabIndex        =   460
            Top             =   300
            Width           =   2295
         End
      End
      Begin VB.Frame fraBlood 
         Height          =   885
         Left            =   120
         TabIndex        =   461
         Top             =   4470
         Width           =   3465
         Begin VB.CheckBox chk 
            Caption         =   "������Ѫ�����������"
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   462
            Top             =   270
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ѫ����ֻ�����м�������ҽʦ���"
            Enabled         =   0   'False
            Height          =   200
            Index           =   85
            Left            =   120
            TabIndex        =   463
            Top             =   570
            Width           =   3230
         End
         Begin VB.CheckBox chk 
            Caption         =   "������Ѫ�ּ�����"
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   464
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame fraCheckDrug 
         Caption         =   "������ҩ�ӿ�"
         Height          =   2160
         Left            =   120
         TabIndex        =   465
         Top             =   120
         Width           =   3465
         Begin VB.CommandButton cmdSet 
            Caption         =   "����"
            Height          =   300
            Left            =   2700
            TabIndex        =   531
            Top             =   0
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҩƷҪ����дԭ��"
            Height          =   240
            Index           =   139
            Left            =   120
            TabIndex        =   487
            Top             =   1200
            Width           =   2940
         End
         Begin VB.OptionButton optPASSVer 
            Caption         =   "����4.0"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   466
            Top             =   1740
            Width           =   975
         End
         Begin VB.OptionButton optPASSVer 
            Caption         =   "����3.0"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   467
            Top             =   1740
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����´�Ժ��ִ�еĽ���ҩƷҽ��"
            Height          =   240
            Index           =   77
            Left            =   120
            TabIndex        =   469
            Top             =   666
            Width           =   3000
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ֹ�´ﳬ����ҩƷҽ��"
            Height          =   240
            Index           =   63
            Left            =   120
            TabIndex        =   470
            Top             =   930
            Width           =   2940
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����´����ҩƷҽ��"
            Height          =   240
            Index           =   65
            Left            =   120
            TabIndex        =   471
            Top             =   400
            Width           =   2100
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   20
            ItemData        =   "frmParClinic.frx":2FF9C
            Left            =   1230
            List            =   "frmParClinic.frx":2FF9E
            Style           =   2  'Dropdown List
            TabIndex        =   472
            Top             =   0
            Width           =   1410
         End
         Begin VB.ComboBox cbo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   27
            ItemData        =   "frmParClinic.frx":2FFA0
            Left            =   1260
            List            =   "frmParClinic.frx":2FFA2
            Style           =   2  'Dropdown List
            TabIndex        =   473
            Top             =   1755
            Width           =   1770
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʹ��ϵͳ����"
            Height          =   240
            Index           =   89
            Left            =   120
            TabIndex        =   468
            Top             =   1440
            Width           =   2940
         End
         Begin VB.CheckBox chk 
            Caption         =   "���ýӿڵ�����־"
            Height          =   240
            Index           =   88
            Left            =   120
            TabIndex        =   103
            Top             =   1440
            Width           =   2940
         End
         Begin VB.Label lblPassVer 
            Caption         =   "��ǰ�汾��"
            Height          =   255
            Left            =   120
            TabIndex        =   474
            Top             =   1740
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������Դ"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   104
            Top             =   1815
            Width           =   1080
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "�������뵥�����ʹ�����뵥�´�ҽ����"
         Height          =   180
         Left            =   3780
         TabIndex        =   523
         Top             =   5115
         Width           =   3240
      End
      Begin VB.Label Label2 
         Caption         =   "סԺҽ��վ�����б�          ��ʾ"
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   5760
         Width           =   3375
      End
      Begin VB.Label lblBloodPrompt 
         Caption         =   "סԺ��Ѫ����ע������"
         Height          =   255
         Index           =   1
         Left            =   3765
         TabIndex        =   475
         Top             =   6405
         Width           =   2655
      End
      Begin VB.Label lblBloodPrompt 
         Caption         =   "������Ѫ����ע������"
         Height          =   255
         Index           =   0
         Left            =   3765
         TabIndex        =   476
         Top             =   5475
         Width           =   2655
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7335
      Index           =   9
      Left            =   2430
      ScaleHeight     =   7305
      ScaleWidth      =   10065
      TabIndex        =   196
      Top             =   45
      Width           =   10095
      Begin VB.Frame fraEprSign 
         Caption         =   "����ǩ��"
         Height          =   7005
         Left            =   5145
         TabIndex        =   534
         Top             =   120
         Width           =   4740
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   270
            Index           =   10
            Left            =   3915
            TabIndex        =   537
            Text            =   "50"
            Top             =   315
            Width           =   390
         End
         Begin VB.CheckBox chk 
            Caption         =   "ǩ��ʹ��ͼƬʱ��ԭͼ"
            Height          =   195
            Index           =   98
            Left            =   75
            TabIndex        =   536
            Top             =   345
            Width           =   2235
         End
         Begin VB.CheckBox chk 
            Caption         =   "ǩ��ʹ��ͼƬ"
            Height          =   195
            Index           =   97
            Left            =   2460
            TabIndex        =   535
            Top             =   795
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfDepartSign 
            Height          =   5805
            Left            =   90
            TabIndex        =   538
            Top             =   1080
            Width           =   4575
            _cx             =   8070
            _cy             =   10239
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParClinic.frx":2FFA4
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            Editable        =   2
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
            Begin VB.CommandButton cmdDepartSelect 
               Caption         =   "��"
               Height          =   240
               Left            =   4260
               TabIndex        =   539
               Top             =   255
               Visible         =   0   'False
               Width           =   300
            End
         End
         Begin VB.Label lblEpr 
            AutoSize        =   -1  'True
            Caption         =   "ʹ��ͼƬǩ����������"
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   541
            Top             =   810
            Width           =   1800
         End
         Begin VB.Label lblEpr 
            AutoSize        =   -1  'True
            Caption         =   "ָ��ǩ��ͼƬ�߶�     ����"
            Height          =   180
            Index           =   2
            Left            =   2430
            TabIndex        =   540
            Top             =   345
            Width           =   2250
         End
      End
      Begin VB.Frame fraEprWrite 
         Caption         =   "������д"
         Height          =   3150
         Left            =   255
         TabIndex        =   225
         Top             =   120
         Width           =   4650
         Begin VB.CheckBox chk 
            Caption         =   "��д��ϵ����Ĵ�Ⱦ�����濨��ǿ����д"
            Height          =   285
            Index           =   182
            Left            =   75
            TabIndex        =   644
            Top             =   2700
            Width           =   3660
         End
         Begin VB.CheckBox chk 
            Caption         =   "������/��Ժ���ʱͬ��������ҳ"
            Height          =   285
            Index           =   99
            Left            =   75
            TabIndex        =   232
            Top             =   2310
            Width           =   3420
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   2
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   231
            Top             =   1853
            Width           =   2565
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ǩ��������Ϊǰ׺����(&P)"
            Height          =   225
            Index           =   95
            Left            =   75
            TabIndex        =   229
            Top             =   1124
            Width           =   2565
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ʾ��ǩλ��(&H)"
            Height          =   240
            Index           =   96
            Left            =   75
            TabIndex        =   230
            Top             =   1506
            Width           =   1695
         End
         Begin VB.OptionButton optSign 
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   1095
            TabIndex        =   227
            Top             =   330
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.OptionButton optSign 
            Caption         =   "ǩ��"
            Height          =   180
            Index           =   1
            Left            =   2535
            TabIndex        =   228
            Top             =   330
            Width           =   915
         End
         Begin VB.CheckBox chk 
            Caption         =   "ǩ����λ(&S)"
            Height          =   285
            Index           =   94
            Left            =   75
            TabIndex        =   226
            Top             =   682
            Width           =   1305
         End
         Begin VB.Label lblEpr 
            AutoSize        =   -1  'True
            Caption         =   "ǩ��ʱ��(&T)"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   234
            Top             =   1913
            Width           =   990
         End
         Begin VB.Label lblEpr 
            Caption         =   "ǩ����ʾ"
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   233
            Top             =   315
            Width           =   870
         End
      End
      Begin VB.Frame fraEprIn 
         Caption         =   "סԺ����"
         Height          =   3705
         Index           =   0
         Left            =   255
         TabIndex        =   210
         Top             =   3420
         Width           =   4650
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   180
            Index           =   12
            Left            =   2940
            TabIndex        =   221
            Text            =   "7"
            Top             =   2805
            Width           =   360
         End
         Begin VB.OptionButton optEprRead 
            Caption         =   "������Ԥ����ѡ��һ���ļ���һ�Ρ�"
            Height          =   195
            Index           =   1
            Left            =   345
            TabIndex        =   220
            Top             =   2470
            Width           =   4260
         End
         Begin VB.OptionButton optEprRead 
            Caption         =   "����Ԥ�����״ζ�ȡȫ��������������ֻ��λ��"
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   219
            Top             =   2150
            Value           =   -1  'True
            Width           =   4260
         End
         Begin VB.Frame fraEprsplit 
            Height          =   30
            Left            =   0
            TabIndex        =   218
            Top             =   1725
            Width           =   4635
         End
         Begin VB.CheckBox chk 
            Caption         =   "��������������д��������"
            Height          =   180
            Index           =   102
            Left            =   75
            TabIndex        =   217
            Top             =   1080
            Width           =   3720
         End
         Begin VB.CheckBox chk 
            Caption         =   "�Զ���ʾ�������"
            Height          =   180
            Index           =   101
            Left            =   75
            TabIndex        =   216
            Top             =   705
            Width           =   3720
         End
         Begin VB.CheckBox chk 
            Caption         =   "(ת�ƺ�Ҫ����д)�Ĺ���������һҳ��ӡ"
            Height          =   330
            Index           =   100
            Left            =   75
            TabIndex        =   215
            Top             =   300
            Width           =   3720
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   11
            Left            =   2595
            TabIndex        =   214
            Text            =   "13"
            Top             =   1440
            Width           =   360
         End
         Begin VB.OptionButton optEprRead 
            Caption         =   "����Ԥ������ȡѡ���ļ�ǰ��    ��Ĺ�������"
            Height          =   195
            Index           =   2
            Left            =   345
            TabIndex        =   222
            Top             =   2790
            Width           =   4260
         End
         Begin VB.Line lnEpr 
            Index           =   0
            X1              =   2595
            X2              =   2985
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label lblEpr 
            Caption         =   "�б�����������󣬹�������    ���Զ��۵�"
            Height          =   240
            Index           =   4
            Left            =   75
            TabIndex        =   224
            Top             =   1440
            Width           =   3780
         End
         Begin VB.Line lnEpr 
            Index           =   1
            X1              =   2940
            X2              =   3330
            Y1              =   2985
            Y2              =   2985
         End
         Begin VB.Label lblEpr 
            Caption         =   "����������Ԥ��"
            Height          =   180
            Index           =   5
            Left            =   75
            TabIndex        =   223
            Top             =   1845
            Width           =   1650
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7815
      Index           =   8
      Left            =   2400
      ScaleHeight     =   7785
      ScaleWidth      =   10065
      TabIndex        =   155
      Top             =   30
      Width           =   10095
      Begin VB.CommandButton cmdPatiSign 
         Caption         =   "���˱������(&P)"
         Height          =   350
         Left            =   6555
         TabIndex        =   566
         Top             =   195
         Width           =   1680
      End
      Begin VB.Frame Frame11 
         Caption         =   "ҽ������վ"
         Height          =   2220
         Left            =   120
         TabIndex        =   209
         Top             =   5355
         Width           =   6240
         Begin VB.CheckBox chk 
            Caption         =   "Ѫ͸����д�°滤���¼"
            Height          =   195
            Index           =   150
            Left            =   240
            TabIndex        =   518
            Top             =   1320
            Width           =   2520
         End
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   255
            Left            =   1680
            TabIndex        =   513
            Top             =   1690
            Width           =   2895
            Begin VB.OptionButton opt���˹��� 
               Caption         =   "ִ��ʱ��"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   515
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton opt���˹��� 
               Caption         =   "����ʱ��"
               Height          =   180
               Index           =   1
               Left            =   1200
               TabIndex        =   514
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "��дƤ�Խ��ʱ��֤���"
            Height          =   195
            Index           =   92
            Left            =   240
            TabIndex        =   212
            Top             =   660
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "����δ�շѲ������ִ��"
            Height          =   195
            Index           =   91
            Left            =   240
            TabIndex        =   211
            Top             =   360
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ִ�б���ʱ�շѻ�������"
            Height          =   195
            Index           =   93
            Left            =   240
            TabIndex        =   213
            Top             =   975
            Width           =   2520
         End
         Begin VB.Label lbl���˹��� 
            Caption         =   "���˹���������"
            Height          =   255
            Left            =   240
            TabIndex        =   516
            Top             =   1680
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "����ҽ��վ"
         Height          =   1920
         Left            =   120
         TabIndex        =   178
         Top             =   120
         Width           =   6240
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   4
            Left            =   4725
            TabIndex        =   185
            Text            =   "0"
            Top             =   390
            Width           =   525
         End
         Begin VB.OptionButton opt������� 
            Caption         =   "��ʾ"
            Height          =   240
            Index           =   2
            Left            =   3060
            TabIndex        =   183
            Top             =   375
            Width           =   750
         End
         Begin VB.OptionButton opt������� 
            Caption         =   "��ֹ"
            Height          =   240
            Index           =   1
            Left            =   2310
            TabIndex        =   182
            Top             =   375
            Width           =   855
         End
         Begin VB.OptionButton opt������� 
            Caption         =   "����ֹ"
            Height          =   240
            Index           =   0
            Left            =   1425
            TabIndex        =   181
            Top             =   375
            Width           =   870
         End
         Begin VB.CheckBox chk 
            Caption         =   "ֻ�����Ѿ�����Ĳ���"
            Height          =   195
            Index           =   78
            Left            =   240
            TabIndex        =   190
            Top             =   1611
            Width           =   2100
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   885
            MaxLength       =   4
            TabIndex        =   187
            Text            =   "180"
            Top             =   690
            Width           =   465
         End
         Begin VB.Frame fraLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   885
            TabIndex        =   179
            Top             =   870
            Width           =   465
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҽ�������������ƺ����ﲡ��"
            Height          =   180
            Index           =   72
            Left            =   240
            TabIndex        =   188
            Top             =   1020
            Value           =   1  'Checked
            Width           =   2685
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҽ���������к�������ڶ����н���"
            Height          =   195
            Index           =   73
            Left            =   240
            TabIndex        =   189
            Top             =   1308
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.Label lblReceptionTime 
            AutoSize        =   -1  'True
            Caption         =   "������ǰ       ���ӽ���"
            Height          =   180
            Left            =   3915
            TabIndex        =   184
            Top             =   405
            Width           =   2070
         End
         Begin VB.Line line 
            X1              =   4635
            X2              =   5340
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblReceptionMode 
            Caption         =   "���˽������"
            Height          =   270
            Left            =   240
            TabIndex        =   180
            Top             =   380
            Width           =   1185
         End
         Begin VB.Label lblRefresh 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ��      ���Զ�ˢ�º���/ת�ﲡ���嵥"
            Height          =   180
            Left            =   480
            TabIndex        =   186
            Top             =   705
            Width           =   3330
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "סԺ��ʿվ"
         Height          =   1635
         Left            =   120
         TabIndex        =   195
         Top             =   3465
         Width           =   6240
         Begin VB.CommandButton cmdLink 
            Caption         =   "��֤"
            Height          =   300
            Left            =   4530
            TabIndex        =   208
            TabStop         =   0   'False
            Top             =   1185
            Width           =   570
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   22
            Left            =   3870
            MaxLength       =   4
            TabIndex        =   207
            Top             =   1200
            Width           =   600
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   21
            Left            =   2085
            MaxLength       =   15
            TabIndex        =   205
            Top             =   1185
            Width           =   1305
         End
         Begin VB.OptionButton optNewCard 
            Caption         =   "˳���+��λ��"
            Height          =   180
            Index           =   0
            Left            =   1860
            TabIndex        =   202
            Top             =   885
            Width           =   1515
         End
         Begin VB.OptionButton optNewCard 
            Caption         =   "˳���+��λ���Ʊ��+��λ��"
            Height          =   180
            Index           =   1
            Left            =   3435
            TabIndex        =   203
            Top             =   885
            Value           =   -1  'True
            Width           =   2715
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ƭ����������������(�°�ר��)"
            Height          =   195
            Index           =   133
            Left            =   240
            TabIndex        =   200
            Top             =   555
            Width           =   3900
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   1
            Left            =   900
            TabIndex        =   199
            Top             =   420
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   915
            MaxLength       =   3
            TabIndex        =   198
            Text            =   "1"
            Top             =   240
            Width           =   300
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "�˿�"
            Height          =   180
            Left            =   3465
            TabIndex        =   206
            Top             =   1230
            Width           =   360
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "���廤�������IP��ַ"
            Height          =   180
            Left            =   240
            TabIndex        =   204
            Top             =   1230
            Width           =   1800
         End
         Begin VB.Label Label12 
            Caption         =   "��ʿվ��λ����"
            Height          =   255
            Left            =   240
            TabIndex        =   201
            Top             =   870
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʾ    ���ڵĲ�����鷴����"
            Height          =   180
            Left            =   525
            TabIndex        =   197
            Top             =   255
            Width           =   2520
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "סԺҽ��վ"
         Height          =   1095
         Left            =   120
         TabIndex        =   191
         Top             =   2145
         Width           =   6240
         Begin VB.CheckBox chk 
            Caption         =   "��ӵ��ȫԺ����Ȩ�޵Ĳ����ߣ�����ʾû�д�λ�Ŀ��һ���"
            Height          =   255
            Index           =   68
            Left            =   240
            TabIndex        =   517
            Top             =   600
            Width           =   5640
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   915
            MaxLength       =   3
            TabIndex        =   193
            Text            =   "1"
            Top             =   240
            Width           =   300
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   2
            Left            =   900
            TabIndex        =   194
            Top             =   420
            Width           =   300
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʾ    ���ڵĲ�����鷴����"
            Height          =   180
            Left            =   525
            TabIndex        =   192
            Top             =   255
            Width           =   2520
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7695
      Index           =   6
      Left            =   2400
      ScaleHeight     =   7665
      ScaleWidth      =   10185
      TabIndex        =   136
      Top             =   0
      Width           =   10215
      Begin VB.Frame Frame6 
         Caption         =   "סԺ·�����̿���"
         Height          =   5895
         Left            =   120
         TabIndex        =   595
         Top             =   120
         Width           =   4335
         Begin VB.CheckBox chk 
            Caption         =   "������ϲ���Ϊ�����ٴ�·�����������"
            Height          =   255
            Index           =   179
            Left            =   240
            TabIndex        =   634
            Top             =   5040
            Width           =   3975
         End
         Begin VB.CheckBox chk 
            Caption         =   "������ǰ���������·����Ŀ"
            Height          =   180
            Index           =   47
            Left            =   240
            TabIndex        =   613
            Top             =   2880
            Width           =   3015
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ǰһ�첻���������ɽ����·����Ŀ"
            Height          =   180
            Index           =   46
            Left            =   240
            TabIndex        =   612
            Top             =   2520
            Width           =   3615
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   611
            Text            =   "30"
            Top             =   1080
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "δ����ʱ�������ҽ��������"
            Height          =   180
            Index           =   45
            Left            =   240
            TabIndex        =   610
            Top             =   2160
            Width           =   2775
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   0
            Left            =   3120
            TabIndex        =   609
            Top             =   1620
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   3120
            MaxLength       =   1
            TabIndex        =   608
            Text            =   "1"
            Top             =   1440
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҽ�������´�ҽ������·�����ϼ�¼"
            Height          =   180
            Index           =   44
            Left            =   240
            TabIndex        =   607
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   2925
            TabIndex        =   606
            Top             =   1275
            Width           =   420
         End
         Begin VB.Frame fraPathExe 
            Height          =   615
            Left            =   240
            TabIndex        =   602
            Top             =   360
            Width           =   3135
            Begin VB.CheckBox chk 
               Caption         =   "��ʿ"
               Height          =   255
               Index           =   43
               Left            =   1320
               TabIndex        =   605
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "ҽ��"
               Height          =   255
               Index           =   42
               Left            =   480
               TabIndex        =   604
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "����·��ִ�л���"
               Height          =   180
               Index           =   41
               Left            =   240
               TabIndex        =   603
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "ƥ��ʱ��Ч��ͬ��·������Ŀ"
            Height          =   180
            Index           =   84
            Left            =   240
            TabIndex        =   601
            Top             =   3240
            Width           =   3015
         End
         Begin VB.CommandButton cmdPathSortSet 
            Caption         =   "·����Ŀ����˳������(&S)"
            Height          =   350
            Left            =   480
            TabIndex        =   600
            Top             =   5400
            Width           =   2415
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ժ������ȡ�����·��"
            Height          =   180
            Index           =   153
            Left            =   240
            TabIndex        =   599
            Top             =   4320
            Width           =   3015
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷҽ����ƥ��Ϊ·������Ŀ"
            Height          =   180
            Index           =   57
            Left            =   240
            TabIndex        =   598
            Top             =   3600
            Width           =   3015
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҩ���ơ�ҽ���˫���ģʽ"
            Height          =   255
            Index           =   58
            Left            =   240
            TabIndex        =   597
            Top             =   4680
            Width           =   3255
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷҽ����ͬ���಻��·������Ŀ"
            Height          =   180
            Index           =   134
            Left            =   240
            TabIndex        =   596
            Top             =   3960
            Width           =   3015
         End
         Begin VB.Label lblSendAdvice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����·��ʱ��ҽ����������ǰʱ��    ��"
            Height          =   180
            Left            =   240
            TabIndex        =   615
            Top             =   1440
            Width           =   3420
         End
         Begin VB.Label lbl��ҩζ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ�䷽�����޸ĵ���ҩζ������    %"
            Height          =   180
            Left            =   240
            TabIndex        =   614
            Top             =   1080
            Width           =   3150
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "����·�����̿���"
         Height          =   5535
         Left            =   5040
         TabIndex        =   587
         Top             =   120
         Width           =   4335
         Begin VB.Frame Frame17 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   2925
            TabIndex        =   592
            Top             =   960
            Width           =   420
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   23
            Left            =   3120
            MaxLength       =   1
            TabIndex        =   591
            Text            =   "1"
            Top             =   1080
            Width           =   300
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   7
            Left            =   3120
            TabIndex        =   590
            Top             =   1260
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   24
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   589
            Text            =   "30"
            Top             =   720
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "����·��ִ�л���"
            Height          =   180
            Index           =   169
            Left            =   240
            TabIndex        =   588
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ�䷽�����޸ĵ���ҩζ������    %"
            Height          =   180
            Left            =   240
            TabIndex        =   594
            Top             =   720
            Width           =   3150
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����·��ʱ��ҽ����������ǰʱ��    ��"
            Height          =   180
            Left            =   240
            TabIndex        =   593
            Top             =   1080
            Width           =   3420
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "סԺ·����ӡ����"
         Height          =   1455
         Left            =   120
         TabIndex        =   577
         Top             =   6120
         Width           =   4335
         Begin VB.Frame Frame12 
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   255
            Left            =   1920
            TabIndex        =   627
            Top             =   1080
            Width           =   2295
            Begin VB.OptionButton optPrintWay 
               Caption         =   "����ʽ"
               Height          =   180
               Index           =   1
               Left            =   1200
               TabIndex        =   629
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optPrintWay 
               Caption         =   "���ʽ"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   628
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.Frame fraRule 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1800
            TabIndex        =   582
            Top             =   360
            Width           =   2415
            Begin VB.OptionButton opt·����ӡ���� 
               Caption         =   "���׶δ�ӡ"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   584
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton opt·����ӡ���� 
               Caption         =   "�����ӡ"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   583
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.OptionButton optÿҳ·����ӡ���� 
            Caption         =   "3��"
            Height          =   180
            Index           =   3
            Left            =   2880
            TabIndex        =   581
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton optÿҳ·����ӡ���� 
            Caption         =   "2��"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   580
            Top             =   -9000
            Width           =   615
         End
         Begin VB.OptionButton optÿҳ·����ӡ���� 
            Caption         =   "2��"
            Height          =   180
            Index           =   1
            Left            =   3480
            TabIndex        =   579
            Top             =   -9000
            Width           =   615
         End
         Begin VB.OptionButton optÿҳ·����ӡ���� 
            Caption         =   "2��"
            Height          =   180
            Index           =   2
            Left            =   2280
            TabIndex        =   578
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "·������ӡ��ʽ��"
            Height          =   180
            Left            =   240
            TabIndex        =   626
            Top             =   1080
            Width           =   1620
         End
         Begin VB.Label lblPrtRule 
            Caption         =   "·������ӡ����"
            Height          =   255
            Left            =   240
            TabIndex        =   586
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblPrintDays 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "·����ÿҳ��ӡ������"
            Height          =   180
            Left            =   240
            TabIndex        =   585
            Top             =   720
            Width           =   1980
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "����·����ӡ����"
         Height          =   1455
         Left            =   5040
         TabIndex        =   567
         Top             =   5760
         Width           =   4335
         Begin VB.OptionButton opt����·����ӡ���� 
            Caption         =   "2��"
            Height          =   180
            Index           =   1
            Left            =   -9000
            TabIndex        =   617
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton opt����·����ӡ���� 
            Caption         =   "2��"
            Height          =   180
            Index           =   0
            Left            =   -9000
            TabIndex        =   616
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optÿҳ·����ӡ���� 
            Caption         =   "2��"
            Height          =   180
            Index           =   4
            Left            =   3480
            TabIndex        =   568
            Top             =   -9000
            Width           =   615
         End
         Begin VB.Frame Frame18 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1800
            TabIndex        =   572
            Top             =   360
            Width           =   2415
            Begin VB.OptionButton opt����·����ӡ���� 
               Caption         =   "���׶δ�ӡ"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   574
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton opt����·����ӡ���� 
               Caption         =   "�����ӡ"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   573
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.OptionButton opt����·����ӡ���� 
            Caption         =   "3��"
            Height          =   180
            Index           =   3
            Left            =   2880
            TabIndex        =   571
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton opt����·����ӡ���� 
            Caption         =   "2��"
            Height          =   180
            Index           =   2
            Left            =   2280
            TabIndex        =   570
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton optÿҳ·����ӡ���� 
            Caption         =   "2��"
            Height          =   180
            Index           =   5
            Left            =   3600
            TabIndex        =   569
            Top             =   -9000
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "·������ӡ����"
            Height          =   255
            Left            =   240
            TabIndex        =   576
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "·����ÿҳ��ӡ������"
            Height          =   180
            Left            =   240
            TabIndex        =   575
            Top             =   720
            Width           =   1980
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   4
      Left            =   2400
      ScaleHeight     =   7425
      ScaleWidth      =   10065
      TabIndex        =   113
      Top             =   15
      Width           =   10095
      Begin VB.Frame fraAdvicePrint 
         Caption         =   "ҽ������ӡ"
         Height          =   2235
         Left            =   4320
         TabIndex        =   141
         Top             =   240
         Width           =   5625
         Begin VB.TextBox txt 
            Height          =   270
            Index           =   26
            Left            =   4050
            MaxLength       =   100
            TabIndex        =   637
            Top             =   1860
            Width           =   555
         End
         Begin VB.TextBox txt 
            Height          =   270
            Index           =   8
            Left            =   2685
            MaxLength       =   100
            TabIndex        =   636
            Top             =   1860
            Width           =   555
         End
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   255
            Index           =   172
            Left            =   2655
            TabIndex        =   631
            Top             =   1215
            Width           =   660
         End
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   255
            Index           =   173
            Left            =   3420
            TabIndex        =   620
            Top             =   915
            Width           =   660
         End
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   255
            Index           =   171
            Left            =   2655
            TabIndex        =   619
            Top             =   915
            Width           =   660
         End
         Begin VB.PictureBox picOptFra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   1
            Left            =   2595
            ScaleHeight     =   195
            ScaleWidth      =   2895
            TabIndex        =   555
            Top             =   300
            Width           =   2900
            Begin VB.OptionButton optPrintDruUse 
               Caption         =   "һ����ҩ��ӡ"
               Height          =   180
               Index           =   2
               Left            =   1545
               TabIndex        =   558
               Top             =   0
               Width           =   1400
            End
            Begin VB.OptionButton optPrintDruUse 
               Caption         =   "��ӡ"
               Height          =   180
               Index           =   1
               Left            =   870
               TabIndex        =   557
               Top             =   0
               Width           =   720
            End
            Begin VB.OptionButton optPrintDruUse 
               Caption         =   "����ӡ"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   556
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.PictureBox picOptFra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   0
            Left            =   2715
            ScaleHeight     =   195
            ScaleWidth      =   2700
            TabIndex        =   551
            Top             =   615
            Width           =   2700
            Begin VB.OptionButton optת��������Ժҽ�� 
               Caption         =   "��������"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   2
               Left            =   1725
               TabIndex        =   552
               Top             =   0
               Width           =   1150
            End
            Begin VB.OptionButton optת��������Ժҽ�� 
               Caption         =   "������"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   855
               TabIndex        =   553
               Top             =   0
               Width           =   900
            End
            Begin VB.OptionButton optת��������Ժҽ�� 
               Caption         =   "������"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   554
               Top             =   0
               Value           =   -1  'True
               Width           =   900
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "ת��"
            Height          =   285
            Index           =   162
            Left            =   1920
            TabIndex        =   547
            Top             =   1215
            Width           =   660
         End
         Begin VB.CheckBox chk 
            Caption         =   "ת��"
            Height          =   255
            Index           =   59
            Left            =   1920
            TabIndex        =   143
            Top             =   915
            Width           =   660
         End
         Begin VB.CheckBox chk 
            Caption         =   "������"
            Height          =   285
            Index           =   60
            Left            =   3420
            TabIndex        =   142
            Top             =   1515
            Width           =   850
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "��ҩҽ��������ʾ����  ������         ������"
            Height          =   180
            Left            =   150
            TabIndex        =   635
            Top             =   1875
            Width           =   3870
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "ת�ƻ�ҳ�������д�ӡ""�ؿ�ҽ��""����"
            Height          =   180
            Left            =   135
            TabIndex        =   633
            Top             =   1560
            Width           =   3060
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "����������һҳ��ӡ"
            Height          =   180
            Left            =   135
            TabIndex        =   632
            Top             =   1245
            Width           =   1620
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "����������һҳ��ӡ"
            Height          =   180
            Left            =   135
            TabIndex        =   618
            Top             =   930
            Width           =   1620
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ����ҩ�÷�������ӡһ��"
            Height          =   180
            Left            =   135
            TabIndex        =   550
            Top             =   285
            Width           =   2340
         End
         Begin VB.Label Label1 
            Caption         =   "ת�ơ���Ժ������ҽ����ӡλ��"
            Height          =   255
            Left            =   135
            TabIndex        =   144
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   "��ҪҩƷ����Ǽǵĸ�ҩ;��"
         Height          =   4770
         Left            =   4320
         TabIndex        =   137
         Top             =   2520
         Width           =   5625
         Begin VB.ListBox lst 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   4050
            IMEMode         =   3  'DISABLE
            Index           =   7
            ItemData        =   "frmParClinic.frx":3002E
            Left            =   120
            List            =   "frmParClinic.frx":30035
            Style           =   1  'Checkbox
            TabIndex        =   140
            Top             =   240
            Width           =   5355
         End
         Begin VB.CommandButton cmdҩƷ�����ҩ;�� 
            Caption         =   "ȫѡ"
            Height          =   300
            Index           =   0
            Left            =   3600
            TabIndex        =   139
            Top             =   4360
            Width           =   900
         End
         Begin VB.CommandButton cmdҩƷ�����ҩ;�� 
            Caption         =   "ȫ��"
            Height          =   300
            Index           =   1
            Left            =   4560
            TabIndex        =   138
            Top             =   4360
            Width           =   900
         End
      End
      Begin VB.Frame fraִ�� 
         Caption         =   " ҽ��ִ�� "
         Height          =   2235
         Left            =   240
         TabIndex        =   83
         Top             =   240
         Width           =   3975
         Begin VB.CheckBox chk 
            Caption         =   "Ƥ��"
            Height          =   200
            Index           =   61
            Left            =   825
            TabIndex        =   502
            Top             =   280
            Width           =   690
         End
         Begin VB.CheckBox chk 
            Caption         =   "����Աֻ���Ա�����ݵǼ�"
            Height          =   210
            Index           =   32
            Left            =   120
            TabIndex        =   500
            Top             =   1560
            Width           =   2520
         End
         Begin VB.CheckBox chk 
            Caption         =   "ִ�е���ӡʱ������˻�ҳ��ӡ"
            Height          =   195
            Index           =   55
            Left            =   120
            TabIndex        =   148
            Top             =   1275
            Width           =   3540
         End
         Begin VB.CheckBox chk 
            Caption         =   "ָ��ҽ������������ִ��"
            Height          =   240
            Index           =   62
            Left            =   120
            TabIndex        =   88
            Top             =   945
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ѫ"
            Height          =   200
            Index           =   74
            Left            =   120
            TabIndex        =   84
            Top             =   280
            Width           =   690
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1150
            MaxLength       =   3
            TabIndex        =   86
            Text            =   "999"
            Top             =   555
            Width           =   510
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ȡ��"
            Height          =   200
            Index           =   87
            Left            =   120
            TabIndex        =   85
            Top             =   585
            Width           =   1035
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   0
            Left            =   1680
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   555
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(0)"
            BuddyDispid     =   196669
            BuddyIndex      =   0
            OrigLeft        =   2400
            OrigTop         =   1380
            OrigRight       =   2655
            OrigBottom      =   1680
            Max             =   999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   0   'False
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "ҽ��ִ�к���Ҫ�˶�"
            Height          =   180
            Left            =   1665
            TabIndex        =   501
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ڵ�ִ�в���"
            Height          =   180
            Index           =   1
            Left            =   1995
            TabIndex        =   117
            Top             =   600
            Width           =   1260
         End
      End
      Begin VB.Frame fra�����ջ� 
         Caption         =   " �����ջ� "
         Height          =   4770
         Left            =   240
         TabIndex        =   89
         Top             =   2520
         Width           =   3975
         Begin VB.CheckBox chk 
            Caption         =   "�ջ�ʱҽ���б�ֻ��ʾ��ǰ������ҽ��"
            Height          =   255
            Index           =   67
            Left            =   210
            TabIndex        =   145
            Top             =   1200
            Width           =   3615
         End
         Begin VB.OptionButton opt���ڷ����ջ� 
            Caption         =   "��������"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   92
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt���ڷ����ջ� 
            Caption         =   "��������"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   91
            Top             =   300
            Width           =   1095
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����ջ�ʱ�Զ���˱���ִ�е���������"
            Height          =   195
            Index           =   36
            Left            =   210
            TabIndex        =   93
            Top             =   600
            Width           =   3680
         End
         Begin VB.CheckBox chk 
            Caption         =   "ȷ��ֹͣ���Զ�ִ�г����ջ�"
            Height          =   195
            Index           =   37
            Left            =   210
            TabIndex        =   94
            Top             =   900
            Width           =   3180
         End
         Begin VB.ListBox lst 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   2790
            IMEMode         =   3  'DISABLE
            Index           =   6
            ItemData        =   "frmParClinic.frx":3003E
            Left            =   210
            List            =   "frmParClinic.frx":30045
            Style           =   1  'Checkbox
            TabIndex        =   96
            Top             =   1860
            Width           =   3525
         End
         Begin VB.Label lblRoll 
            Caption         =   "�����ջ�ģʽ"
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lblSend 
            Caption         =   "���·�ҩ��ʽ����ҩһ����ҩ�Ͳ��ջ�"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   1590
            Width           =   3255
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   2
      Left            =   2400
      ScaleHeight     =   7425
      ScaleWidth      =   10065
      TabIndex        =   105
      Top             =   0
      Width           =   10095
      Begin VB.Frame fra��� 
         Caption         =   "���﷢���������ʱ��������д"
         Height          =   1530
         Left            =   5160
         TabIndex        =   55
         Top             =   3600
         Width           =   4695
         Begin VB.CommandButton cmd���﷢�ͼ����� 
            Caption         =   "ȫ��"
            Height          =   300
            Index           =   1
            Left            =   3600
            TabIndex        =   58
            Top             =   720
            Width           =   900
         End
         Begin VB.CommandButton cmd���﷢�ͼ����� 
            Caption         =   "ȫѡ"
            Height          =   300
            Index           =   0
            Left            =   3600
            TabIndex        =   57
            Top             =   360
            Width           =   900
         End
         Begin VB.ListBox lst 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   56
            Top             =   360
            Width           =   3300
         End
      End
      Begin VB.Frame fra���﷢�� 
         Caption         =   "���﷢��ѡ��"
         Height          =   1455
         Left            =   5160
         TabIndex        =   47
         Top             =   240
         Width           =   4695
         Begin VB.CheckBox chk 
            Caption         =   "����ҽ�����ͺ��������֧��"
            Height          =   180
            Index           =   12
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   3015
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʱ������ִ�е���Ϊ��ִ��"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   2820
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "���﷢�͵������ "
         Height          =   1080
         Index           =   0
         Left            =   5160
         TabIndex        =   50
         Top             =   1920
         Width           =   4695
         Begin VB.OptionButton opt���͵������� 
            Caption         =   "�շѵ���"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   51
            Top             =   330
            Width           =   1020
         End
         Begin VB.OptionButton opt���͵������� 
            Caption         =   "���ʵ���"
            Height          =   180
            Index           =   1
            Left            =   1395
            TabIndex        =   52
            Top             =   330
            Width           =   1020
         End
         Begin VB.OptionButton opt���͵������� 
            Caption         =   "����ʱ��ȷ��"
            Height          =   180
            Index           =   2
            Left            =   2565
            TabIndex        =   53
            Top             =   330
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.CheckBox chk 
            Caption         =   "ֻ�к�Լ��λ���˵�ҽ���ſ��Է���Ϊ���ʵ�"
            Height          =   195
            Index           =   6
            Left            =   255
            TabIndex        =   54
            Top             =   630
            Width           =   3960
         End
      End
      Begin VB.Frame fraSendNO 
         Caption         =   "���﷢�͵��ݲ�������"
         Height          =   3435
         Left            =   240
         TabIndex        =   37
         Top             =   3600
         Width           =   4695
         Begin VB.CheckBox chk 
            Caption         =   "����ҽ������ʱһ����鷢��Ϊһ�ŵ���"
            Height          =   180
            Index           =   147
            Left            =   465
            TabIndex        =   510
            Top             =   2760
            Width           =   4080
         End
         Begin VB.CommandButton cmd���﷢��һ�ŵ������ 
            Caption         =   "ȫ��"
            Height          =   300
            Index           =   1
            Left            =   3600
            TabIndex        =   45
            Top             =   2160
            Width           =   900
         End
         Begin VB.CommandButton cmd���﷢��һ�ŵ������ 
            Caption         =   "ȫѡ"
            Height          =   300
            Index           =   0
            Left            =   3600
            TabIndex        =   44
            Top             =   1800
            Width           =   900
         End
         Begin VB.ListBox lst 
            Columns         =   3
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   480
            Style           =   1  'Checkbox
            TabIndex        =   43
            Top             =   1815
            Width           =   2940
         End
         Begin VB.OptionButton opt���͵��ݹ��� 
            Caption         =   "����ͬһ���ҽ����ִͬ�п���ֻ����һ�ŵ���"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   42
            Top             =   1515
            Width           =   4140
         End
         Begin VB.OptionButton opt���͵��ݹ��� 
            Caption         =   "ÿ�η���ҽ��ֻ����һ�ŵ���"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   40
            Top             =   900
            Width           =   3060
         End
         Begin VB.CheckBox chk 
            Caption         =   "һ����ҩ�ļ�ʹ�����㲻ͬҲ����Ϊһ�ŵ���"
            Height          =   255
            Index           =   9
            Left            =   465
            TabIndex        =   46
            Top             =   3045
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin VB.OptionButton opt���͵��ݹ��� 
            Caption         =   "�������ҽ������ִͬ�п���ֻ����һ�ŵ���"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   41
            Top             =   1200
            Width           =   4140
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ͬ��ϵ�ҽ���ֱ��������"
            Height          =   180
            Index           =   7
            Left            =   240
            TabIndex        =   38
            Top             =   315
            Width           =   2760
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ʼʱ�䲻��ͬһ��ķֱ��������"
            Height          =   180
            Index           =   8
            Left            =   240
            TabIndex        =   39
            Top             =   600
            Width           =   3480
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   "����Ϊ���ʻ��۵����������"
         Height          =   2775
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   4695
         Begin VB.CheckBox chk 
            Caption         =   "����ҽ������ʱ������������"
            Height          =   200
            Index           =   54
            Left            =   120
            TabIndex        =   36
            Top             =   2160
            Width           =   2640
         End
         Begin VB.CommandButton cmd���ͻ������ 
            Caption         =   "ȫѡ"
            Height          =   300
            Index           =   0
            Left            =   3600
            TabIndex        =   34
            Top             =   600
            Width           =   900
         End
         Begin VB.CommandButton cmd���ͻ������ 
            Caption         =   "ȫ��"
            Height          =   300
            Index           =   1
            Left            =   3600
            TabIndex        =   35
            Top             =   960
            Width           =   900
         End
         Begin TabDlg.SSTab SendPriceType 
            Height          =   1665
            Left            =   120
            TabIndex        =   31
            Top             =   255
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2937
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabMaxWidth     =   882
            TabCaption(0)   =   "����"
            TabPicture(0)   =   "frmParClinic.frx":3004E
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "lst(4)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "סԺ"
            TabPicture(1)   =   "frmParClinic.frx":3006A
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "lst(3)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.ListBox lst 
               Columns         =   3
               Height          =   1110
               IMEMode         =   3  'DISABLE
               Index           =   3
               ItemData        =   "frmParClinic.frx":30086
               Left            =   75
               List            =   "frmParClinic.frx":30088
               Style           =   1  'Checkbox
               TabIndex        =   33
               Top             =   360
               Width           =   3180
            End
            Begin VB.ListBox lst 
               Columns         =   3
               Height          =   1110
               IMEMode         =   3  'DISABLE
               Index           =   4
               ItemData        =   "frmParClinic.frx":3008A
               Left            =   -74925
               List            =   "frmParClinic.frx":3008C
               Style           =   1  'Checkbox
               TabIndex        =   32
               Top             =   360
               Width           =   3180
            End
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7335
      Index           =   10
      Left            =   2565
      ScaleHeight     =   7305
      ScaleWidth      =   10065
      TabIndex        =   235
      Top             =   30
      Width           =   10095
      Begin VB.Frame fraEprControl 
         Caption         =   "���Ӳ���"
         Height          =   7080
         Left            =   120
         TabIndex        =   236
         Top             =   150
         Width           =   9855
         Begin VB.CheckBox chk 
            Caption         =   "���������ȱ�Ŀ������"
            Height          =   285
            Index           =   149
            Left            =   7635
            TabIndex        =   512
            Top             =   2093
            Width           =   2100
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����Զ�д�벡����ҳ"
            Height          =   285
            Index           =   148
            Left            =   5265
            TabIndex        =   511
            Top             =   2093
            Width           =   2130
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   15
            Left            =   1815
            TabIndex        =   246
            Text            =   "7"
            Top             =   1800
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   16
            Left            =   1815
            TabIndex        =   245
            Text            =   "10"
            Top             =   2160
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   14
            Left            =   540
            TabIndex        =   244
            Text            =   "10"
            Top             =   660
            Width           =   435
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   13
            Left            =   2325
            TabIndex        =   243
            Text            =   "7"
            Top             =   315
            Width           =   510
         End
         Begin VB.CheckBox chk 
            Caption         =   "����¼�������������"
            Height          =   285
            Index           =   105
            Left            =   5265
            TabIndex        =   242
            Top             =   1748
            Width           =   2100
         End
         Begin VB.CheckBox chk 
            Caption         =   "��������¼�����ԭ��"
            Height          =   285
            Index           =   106
            Left            =   7635
            TabIndex        =   241
            Top             =   1748
            Width           =   2130
         End
         Begin VB.CommandButton cmdEprUp 
            Caption         =   "����(&U)"
            Height          =   350
            Left            =   8565
            TabIndex        =   240
            Top             =   675
            Width           =   1200
         End
         Begin VB.CommandButton cmdEprDown 
            Caption         =   "����(&D)"
            Height          =   350
            Left            =   8565
            TabIndex        =   239
            Top             =   1140
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������������ܹ鵵"
            Height          =   285
            Index           =   103
            Left            =   165
            TabIndex        =   238
            Top             =   1025
            Width           =   2505
         End
         Begin VB.CheckBox chk 
            Caption         =   "��������¼��������"
            Height          =   285
            Index           =   104
            Left            =   165
            TabIndex        =   237
            Top             =   1395
            Width           =   2160
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfEpr 
            Height          =   1395
            Index           =   0
            Left            =   5280
            TabIndex        =   247
            Top             =   210
            Width           =   3180
            _cx             =   5609
            _cy             =   2461
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
         Begin VSFlex8Ctl.VSFlexGrid vsfEpr 
            Height          =   4155
            Index           =   2
            Left            =   5265
            TabIndex        =   248
            Top             =   2820
            Width           =   4500
            _cx             =   7937
            _cy             =   7329
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            Begin VB.CommandButton cmdEprSelect 
               Caption         =   "��"
               Height          =   240
               Index           =   2
               Left            =   4140
               TabIndex        =   258
               Top             =   270
               Visible         =   0   'False
               Width           =   300
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfEpr 
            Height          =   4155
            Index           =   1
            Left            =   165
            TabIndex        =   249
            Top             =   2820
            Width           =   4500
            _cx             =   7937
            _cy             =   7329
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            Begin VB.CommandButton cmdEprSelect 
               Caption         =   "��"
               Height          =   240
               Index           =   1
               Left            =   4125
               TabIndex        =   257
               Top             =   255
               Visible         =   0   'False
               Width           =   300
            End
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���Ӳ������ֿ��ҷ�Χ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   5265
            TabIndex        =   256
            Tag             =   "0"
            Top             =   2520
            Width           =   1800
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���Ӳ��������ҷ�Χ "
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   165
            TabIndex        =   255
            Top             =   2520
            Width           =   1890
         End
         Begin VB.Line lnEpr 
            Index           =   4
            X1              =   1755
            X2              =   2185
            Y1              =   1995
            Y2              =   1995
         End
         Begin VB.Line lnEpr 
            Index           =   5
            X1              =   1755
            X2              =   2185
            Y1              =   2340
            Y2              =   2340
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ������ȱʡΪ     ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   165
            TabIndex        =   254
            Top             =   1800
            Width           =   2250
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ�������Ϊ     ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   12
            Left            =   165
            TabIndex        =   253
            Top             =   2145
            Width           =   2250
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������˳��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   8565
            TabIndex        =   252
            Top             =   210
            Width           =   1080
         End
         Begin VB.Line lnEpr 
            Index           =   2
            X1              =   2340
            X2              =   2875
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Line lnEpr 
            Index           =   3
            X1              =   540
            X2              =   990
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ��     �����Զ�ˢ�µȴ���������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   165
            TabIndex        =   251
            Top             =   670
            Width           =   2970
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�����������ȱʡ����Ϊ      ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   165
            TabIndex        =   250
            Top             =   315
            Width           =   2880
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   11
      Left            =   5640
      ScaleHeight     =   5505
      ScaleWidth      =   6345
      TabIndex        =   259
      TabStop         =   0   'False
      Top             =   600
      Width           =   6375
      Begin VB.CheckBox chk 
         Caption         =   "����δ�շѵ�ҽ��ִ�����"
         Height          =   255
         Index           =   108
         Left            =   360
         TabIndex        =   270
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CheckBox chk 
         Caption         =   "��дƤ�Խ��ʱ��֤���"
         Height          =   255
         Index           =   107
         Left            =   360
         TabIndex        =   269
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   20
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   268
         Top             =   1275
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   19
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   266
         Top             =   915
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   18
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   264
         Top             =   555
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   17
         Left            =   1155
         MaxLength       =   3
         TabIndex        =   263
         Top             =   555
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   9
         Left            =   600
         MaxLength       =   6
         TabIndex        =   261
         Top             =   195
         Width           =   615
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfWaittingMixDept 
         Height          =   2415
         Left            =   360
         TabIndex        =   272
         Top             =   2760
         Width           =   5415
         _cx             =   9551
         _cy             =   4260
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmParClinic.frx":3008E
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
         Editable        =   2
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
         BackStyle       =   0  'Transparent
         Caption         =   "����Һ���̿��ң�"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   271
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ƥ����ǰ       ��������"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   267
         Top             =   1320
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Һ��ǰ       ��������"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   265
         Top             =   960
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ĭ�ϵ���      ����/���ӣ��� Ĭ�ϵ�ϵ��"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   262
         Top             =   600
         Width           =   3420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ        ���Զ�ˢ�²����嵥"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   260
         Top             =   240
         Width           =   2520
      End
   End
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   8205
      Left            =   0
      ScaleHeight     =   8205
      ScaleWidth      =   2415
      TabIndex        =   106
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   2250
         TabIndex        =   108
         Top             =   0
         Width           =   2250
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5250
            Left            =   0
            TabIndex        =   109
            Top             =   720
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   9260
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager imgFunc 
            Left            =   1800
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParClinic.frx":30114
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   110
            Top             =   0
            Width           =   2200
            _Version        =   589884
            _ExtentX        =   3881
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   5820
         Left            =   2280
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5820
         ScaleWidth      =   45
         TabIndex        =   107
         Top             =   120
         Width           =   45
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   6765
         Left            =   0
         TabIndex        =   111
         Top             =   0
         Width           =   2400
         _Version        =   589884
         _ExtentX        =   4233
         _ExtentY        =   11933
         _StockProps     =   64
      End
      Begin XtremeCommandBars.ImageManager imgType 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmParClinic.frx":36010
      End
   End
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   12630
      TabIndex        =   477
      Top             =   8205
      Width           =   12630
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   115
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   102
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   99
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   11400
         TabIndex        =   98
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   10245
         TabIndex        =   97
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   116
         Top             =   165
         Width           =   4095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ҳ���(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   114
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&S)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   101
         Top             =   168
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmParClinic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset '������ؼ���Ӧ��¼����ͬһ���������ܶ�Ӧһ�����ؼ���
Private marrFunc(2) As String
Private mlngPreFind As Long
Private mblnOk As Boolean
Private mobjPass As Object     '������ҩ���ӿ�
Private mblnUseBlood As Boolean    '�Ƿ�װ����Ѫ��

Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk
    chk_����ҩ�������� = 1
    chk_�����´��Ƥ�� = 2
    chk_����ҩ���䵥�� = 3
    chk_����ǼǴ����� = 5
    chk_��Լ��λ���ͼ��� = 6
    chk_��ͬ��Ϸֵ��� = 7
    chk_��ͬ��ķֵ��� = 8
    chk_һ������һ�ŵ��� = 9
    
    chk_���﷢�ͱ����Զ�ִ�� = 10
    chk_����ҽ�����ͺ��������֧�� = 12
    chk_סԺ����ȱʡһ���� = 13
    chk_סԺҩ���䵥�� = 14
    chk_סԺҩ�������� = 15
    chk_�³�Ժҽ������Ժ��� = 16
    chk_����ִ�к�������ҽ�� = 17
    chk_����ס����ҽ���´� = 18
    chk_סԺ�´��Ƥ�� = 19
    chk_סԺ�´��Զ����� = 20
    
    
    chk_סԺ����У�Է��� = 21
    chk_����ҽ�����ͼ��δ��Чҽ�� = 22
    chk_ҩ���������ƽ���ʱ�� = 23
    chk_�����ڷ�ҩ���ͽ���ʱ�� = 25
    chk_У��ȷ��ֹͣ�������ӡ = 26
    chk_У��ȷ��ֹͣ����ǩ�� = 27
    chk_����У�� = 28
    chk_������ͣ = 29
    chk_�Ǽ�Ƥ����֤��� = 30
    chk_������ҽ��ҽ�� = 31
    chk_ʵϰҽ��ͣ������� = 33
    chk_�����ջ��Զ���˱��� = 36
    chk_ȷ��ֹͣ���Զ��ջ� = 37
    chk_סԺ���ͼ��ҽ������ = 38
    chk_סԺ�����Զ�ִ�г��� = 39
    chk_סԺ�����Զ�ִ������ = 40
    chk_������Ҫ����ִ�� = 168
    
    chk_�����Ǽ���Ч���� = 11
    chk_���ﴦ���������� = 52
    
    chk_ҩƷ�������ҽ�� = 4
    chk_����ҽ��������Ч = 24
    chk_�´�ҽ��ʱ��ʾ���� = 66
    chk_һ��������������Ŀ = 34
    chk_���˳�Ժҽ������������Ժ = 81
    chk_�´��Ժҽ���������Ժ = 50
    
    
    chk_����ҽ���������������� = 54
    chk_סԺҩ�����Ͳ�����ҩ�� = 64
    chk_δ����������ֹ��ҩ = 0
        
    chk_���������Һ���Ч�����Ĳ��� = 82
    chk_�ٴ�����վ����ʹ��zlPlugIn���� = 79
    chk_ҽ������ʱ��������ԭ�� = 86
    chk_ҽ����ҩ�������� = 138
    
    chk_��Ѫҽ��ִ�к���Ҫ�˶� = 74
    chk_Ƥ��ҽ��ִ�к���Ҫ�˶� = 61
    
    chk_ҽ��ִ����Ч���� = 87
    chk_ָ��ҽ������������ִ�� = 62
                    
    chk_����Ѫ�����ϵͳ = 135
    chk_��Ѫҽ�����ͺ���ܷ�Ѫ = 176
    chk_��Ѫ���벻��ʾѪҺ��� = 175
    chk_�´���Ѫ����ʱȷ����Ѫ��Ϣ = 181
    chk_ѪҺ���պ������ִ�еǼ� = 183
    
    chk_��Ѫ�ּ����� = 35
    chk_��Ѫ����������� = 53
    chk_��Ѫ���������м�������ҽʦ = 85
        
    
    chk_����ҩ��ּ����� = 75
    chk_����ҩ��ʹ���Ա�ҩ = 76
    chk_��ҽ��С����п���ҩ����� = 137
    
    chk_�����ּ����� = 80
    chk_������Ȩ���� = 49
    chk_�����ּ���� = 140
    chk_����ҽʦ�ﵽ�����ȼ�������� = 83
            
    chk_����ҩ�� = 65
    chk_��ֹ�´ﳬ����ҩƷҽ�� = 63
    chk_����Ժ��ִ�н���ҩƷ = 77
    chk_�ӿڵ�����־_��ͨ = 88   '��ͨ�ӿ���־���� 65522
    chk_ʹ��ϵͳ����_���� = 89   '�����ӿ�ϵͳ���ù��ܿ��Ʋ��� 65198
    chk_����ҩƷҪ����дԭ�� = 139
    '�ٴ�·��Ӧ��
    chk_����·��ִ�л��� = 41
    chk_·��ִ�л���ҽ������ = 42
    chk_·��ִ�л��ڻ�ʿ���� = 43
    chk_·������ҽ��ҽ����ʾ = 44
    chk_δ����ʱ�������ҽ�������� = 45
    chk_����ǰһ�첻���������ɽ���·����Ŀ = 46
    chk_������ǰ���������·����Ŀ = 47
    chk_ƥ��ʱ��Ч��ͬ��·������Ŀ = 84
    chk_ҩƷҽ����ƥ��Ϊ·������Ŀ = 57
    chk_��Ժ������ȡ�����·�� = 153
    chk_����ҩ���ƺ�ҽ���˫���ģʽ = 58
    chk_ҩƷҽ����ͬ���಻��·����ҽ�� = 134
    
    chk_��������·��ִ�л��� = 169

    chk_����ҽ���´�ʱ����ѡ������ʾҩƷ��� = 48
    chk_סԺҽ���´�ʱ����ѡ������ʾҩƷ��� = 51
    chk_ִ�е���ӡʱ������˻�ҳ��ӡ = 55
    chk_������ת�ƻ�ҳ = 59
    chk_ת�ƻ�ҳ�������д�ӡ�ؿ�ҽ�� = 60
    chk_�ջ�ʱҽ���б�ֻ��ʾ��ǰ������ҽ�� = 67
    chk_ʹ����������ʱ�� = 69
    chk_��ҽ���Ҳ�ʹ����Ŀ = 70
    chk_ҽ���ͻ�ʿ�ֱ���д������ҳ = 71
    chk_ҽ�������������ƺ����ﲡ�� = 72
    chk_ҽ���������к�������ڶ����н��� = 73
    chk_ֻ�����Ѿ�����Ĳ��� = 78
    chk_����δ�շѲ������ִ�� = 91
    chk_��дƤ�Խ��ʱ��֤��� = 92
    chk_ִ�б���ʱ�շѻ������� = 93
    chk_Ѫ͸����д�°滤���¼ = 150
    
    chk_ǩ����λ = 94
    chk_ǩ������ǰ� = 95
    chk_��ʾ��ǩλ�� = 96
    chk_ǩ��ʹ��ͼƬ = 97
    chk_ǩ��ʹ��ԭͼ = 98
    chk_�������ͬ����ҳ = 99
    chk_ת�ƺ���д�Ĳ�������һҳ��ӡ = 100
    chk_סԺ�����Զ���ʾ������� = 101
    chk_Ҫ������д�������� = 102
    chk_�������������ܹ鵵 = 103
    chk_��������¼�������� = 104
    chk_����¼������������� = 105
    chk_��������¼�����ԭ�� = 106
    chk_�����Զ�д�벡����ҳ = 148
    chk_���������ȱ�Ŀ������ = 149
    
    chk_������ҺƤ����֤��� = 107
    chk_������Һδ�շ�������� = 108
    
    '�°�סԺ��ʿ����վ
    chk_��Ƭ��������� = 133
    
    chk_����ʾ�޴�λ�Ĳ������� = 68
    chk_������ϲ���Ϊ�����ٴ�·����������� = 179
    
    '����ͼ
    chk_����ͼ��ʾ����ʱ�� = 109
    chk_����ͼģʽ = 110
    chk_��¶�ߵ���ʾλ�� = 111
    chk_����ͼ����ʾ������ = 112
    
    '��¼��
    chk_������ʾ��ʽ = 113
    chk_ֻ�ڵ�ǰҳ��ʾ��ҳ���� = 114
    chk_��Ӧ��ݻ����ļ� = 115
    chk_�ļ�ҳ����� = 116
    chk_ǩ������ʾ��ʽ = 117
    chk_��������ͬ�� = 167
    
    '���µ�
    chk_�Զ���־������� = 118
    chk_�Զ���־����40�̶���С��ʾ = 119
    chk_�Զ���־��˳�������� = 120
    chk_�������ڸ�ʽ = 121
    chk_��Ʊ�־���Զ�תΪ��Ժ = 122
    chk_��������������Ժ = 123
    chk_ȫ�������ʾСʱ = 124
    chk_������Ŀ��ʾ�������� = 125
    chk_����������� = 126
    chk_��ͼ�̶ȵ�����ʾ = 127
    chk_��ʾ�����Ϣ = 128
    chk_�ٴ�����ֹͣǰ�α�ע = 129
    chk_��ӡҽԺ���� = 130
    chk_Ӥ��סԺ������0��ʼ���� = 131
    chk_�������˵����Ϣ = 132
    chk_������Ŀ������ʾ = 180
    chk_С��ȱʡ��ʶ = 188
    
    
    'סԺ��ҳ
    chk_���֤���� = 136
    
    chk_סԺ����ҽ������ʱһ����鷢��Ϊһ�ŵ��� = 141
    chk_����ִ�еǼ� = 32
    
    chk_���뵥���û��ڻ��� = 90
    chk_���뵥���û����������� = 142
    chk_���뵥���û���������Ѫ = 143
    chk_���뵥���û���������� = 144
    chk_���뵥���û��������� = 145
    
    chk_���뵥���û���סԺ���� = 157
    chk_���뵥���û���סԺ��Ѫ = 156
    chk_���뵥���û���סԺ���� = 152
    chk_���뵥���û���סԺ��� = 151
    
    chk_��ƻ��������дҪ�� = 146
    chk_�������ҽ������ʱһ����鷢��Ϊһ�ŵ��� = 147
    chk_��������������ɵ��� = 154
    chk_���¼��ʱ�����Զ���ȡ = 170
    chk_·������ԭ����ֵ����ѡȡ = 155
    chk_ת��ת��ҽ�ƻ�������������¼�� = 134
    chk_�������뵥�����ʹ�����뵥�´�ҽ������ = 158
    chk_�������뵥�����ʹ�����뵥�´�ҽ��סԺ = 159
    chk_����ҩƷ�ֿ����� = 160
    chk_������ɺ�ر�ҽ������ = 161
    chk_ͣ��ʱ¼��ԭ�� = 56
    chk_������ת�ƻ�ҳ = 162
    chk_����ҩ�����հ������������� = 163
    chk_�������ֻ����¼��M��ͷ��������̬ѧ���� = 164
    
    chk_����ҽ������Ƥ������ = 165
    chk_סԺҽ������Ƥ������ = 166
    
    chk_����������ҳ = 171
    chk_������������ҳ = 173
    chk_סԺ�ǼǴ����� = 174
    chk_Ƥ��������ҩ���� = 177
    chk_Ƥ��������ҩסԺ = 178
    
    chk_����������ҳ = 172
    
    chk_��Ⱦ�����濨ǿ����д = 182
    chk_��������´�ҽ���ɻ���������Ҵ��� = 184
    chk_����δ����ҽ��ʱ��ֹ����ת��ҽ�� = 185
    chk_������ҽ������¼����ҽ��� = 186
    chk_����д����� = 187
End Enum

Private Enum constCbo
    cbo_סԺ��ҳ��׼ = 0
    cbo_סԺҽ��վ�б���ʾ = 1
    cbo_����ǩ����ʾʱ�� = 2
    
    '��¼������
    cbo_��ǩģʽ = 3
    cbo_ǩ������ʾģʽ = 13
    cbo_С��ȱʡ��ʶ = 14
    '����ͼ����
    cbo_��������ʾ = 4
    cbo_�쳣����ʾ = 5
    cbo_������״ε����� = 6
    cbo_���������쳣�� = 7
    cbo_��¶�½��쳣�� = 8
    cbo_������־���� = 9
    cbo_������־λ�� = 10
    cbo_��������˳�� = 11
    cbo_��¶�½�˳�� = 12
    '���µ�
    cbo_δ��˵����ʾλ�� = 15
    cbo_���²�����ʾ��ʽ = 16
    cbo_������������ʾλ�� = 17
    cbo_�������������ʾλ�� = 18
    cbo_��������ȱʡ��ʽ = 19
    cbo_��־˵����ʱ�����ӷ��� = 21
    cbo_��Ժ�Զ���־ = 22
    cbo_����Զ���־ = 23
    cbo_ת���Զ���־ = 24
    cbo_�����Զ���־ = 25
    cbo_�����Զ���־ = 30
    cbo_��Ժ�Զ���־ = 31
    cbo_�����Զ���־ = 32
    cbo_�����Զ���־ = 33
    cbo_�����Զ���־ = 34
    cbo_ת�����Զ���־ = 35
    
    cbo_������ҩ�ӿ� = 20
    cbo_��ҩ�䷽ = 26
    cmd_����������Դ = 27
    cbo_����ҩ�����Ҷ��շ��� = 28
    cbo_סԺҩ�����Ҷ��շ��� = 29
    cbo_��Ѫ�ɼ�Ĭ���������� = 36
    cbo_סԺ����ִ���Զ���ɷ��� = 37
End Enum

Private Enum constUpDown
    ud_ҽ��ִ����Ч���� = 0
    ud_���ﴦ���������� = 3
    ud_�����Ǽ���Ч���� = 7
    
    ud_��¼ҽ��ʶ���� = 8
    ud_�����¿�ҽ����� = 11
    ud_��¼������¼������ = 21
    ud_�������߹̶�������� = 22
    ud_����ҹ�࿪ʼʱ�� = 23
    ud_����ҹ�����ʱ�� = 24
    ud_���±��̶�������� = 26
    ud_���¿�ʼʱ�� = 27
    
End Enum

Private Enum constTxtUd
    txtud_����ҹ�࿪ʼʱ�� = 23
    txtud_����ҹ�����ʱ�� = 24
End Enum

Private Enum constTxt
    txt_������Ѫ����ע������ = 0
    txt_סԺ��Ѫ����ע������ = 1
    txt_��ҩ�䷽�����޸ĵ���ҩζ������ = 2
    txt_����ʱҽ����������ǰʱ������ = 3
    txt_������ǰ������� = 4
    txt_����ҽ��վ�����б��Զ�ˢ������ = 5
    txt_סԺҽ������������ = 6
    txt_סԺ��ʿ���������� = 7

    
    txt_ǩ��ʹ��ͼƬ�߶� = 10
    txt_������N���Զ��۵� = 11
    txt_����������Ԥ��N�� = 12
    txt_���Ӳ������ȱʡ���� = 13
    txt_���Ӳ�����������ˢ�¼�� = 14
    txt_���Ӳ�������ȱʡ���� = 15
    txt_���Ӳ�������������� = 16
    
    txt_������Һ�Զ�ˢ�²��� = 9
    txt_������Һ���� = 17
    txt_������Һ��ϵ�� = 18
    txt_������Һ��Һ���� = 19
    txt_������ҺƤ������ = 20
    txt_���廤��IP��ַ = 21
    txt_���廤��IP�˿� = 22
    txt_������ע���� = 25
    txt_���¸��Ժϸ���� = 28
    txt_������ҩ�䷽�����޸ĵ���ҩζ������ = 24
    txt_��������ʱҽ����������ǰʱ������ = 23
    txt_��������ҩҽ��������ʾ���� = 8
    txt_��������ҩҽ��������ʾ���� = 26
End Enum

Private Enum constListBox
    lst_���﷢��һ�ŵ������ = 0
    lst_���﷢�ͼ����� = 1
    lst_סԺ�����Ժ��� = 2
    lst_סԺ���ͻ������ = 3
    lst_���﷢�ͻ������ = 4    '����Ϊ���۵����������
    lst_���ջصķ�ҩ���� = 6
    lst_ҩƷ�����ҩ;�� = 7
    lst_����ִ���Զ����ҽ����� = 8
End Enum

'������ҩ�ӿ�
Private Enum lblEnum
    lbl_����������Դ = 0
End Enum

Private mrs����ҩ������ As ADODB.Recordset
Private mrsסԺҩ������ As ADODB.Recordset
Private mrsסԺִ�ж��� As ADODB.Recordset

Private mcol���� As Collection '������д����˵���Ŀ���
Private mcolStop���� As Collection  '����ͣ��¼��ԭ��Ŀ���
Private mrsAdvice As New ADODB.Recordset '��¼ҽ�����ݶ���

Private Sub cmdAdd_Click(Index As Integer)
    Dim lngIndex As Long, i As Long
    
    If cbo(Index).ListCount > 0 Then
        lngIndex = cbo(Index).ItemData(cbo(Index).ListCount - 1)
    Else
        lngIndex = 0
    End If
    cbo(Index).AddItem "����" & lngIndex + 1
    cbo(Index).ItemData(cbo(Index).NewIndex) = lngIndex + 1
    
    If Index = cbo_����ҩ�����Ҷ��շ��� Then
        mrs����ҩ������.AddNew
        mrs����ҩ������!���� = lngIndex + 1
    ElseIf Index = cbo_סԺҩ�����Ҷ��շ��� Then
        mrsסԺҩ������.AddNew
        mrsסԺҩ������!���� = lngIndex + 1
    ElseIf Index = cbo_סԺ����ִ���Զ���ɷ��� Then
        mrsסԺִ�ж���.AddNew
        mrsסԺִ�ж���!���� = lngIndex + 1
    End If
    cbo(Index).ListIndex = cbo(Index).ListCount - 1
        
    If Index = cbo_סԺ����ִ���Զ���ɷ��� Then
        For i = 0 To lst(lst_����ִ���Զ����ҽ�����).ListCount - 1
            lst(lst_����ִ���Զ����ҽ�����).Selected(i) = False
        Next
        Frame14.Tag = "���޸�"
    Else
        With vsfDrugStore(Index)
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("����")) = 0
                .TextMatrix(i, .ColIndex("ȱʡ")) = ""
                If Index = cbo_����ҩ�����Ҷ��շ��� Then
                    .TextMatrix(i, .ColIndex("��ҩ����")) = "�Զ�����"
                End If
            Next
        End With
        SST.Tag = "���޸�"
    End If
End Sub

Private Sub cmdAddMed_Click()
    frmInMedSetup.ShowMe "", "", "", "����", Me
    cmdModify.Enabled = vsfMecItem.Rows > 1:  cmdDelete.Enabled = vsfMecItem.Rows > 1
End Sub

Private Sub cmdDel_Click(Index As Integer)
    On Error Resume Next
    If cbo(Index).ListCount = 0 Then Exit Sub
    If MsgBox("�Ƿ�ɾ���˷�����", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
        If Index = cbo_����ҩ�����Ҷ��շ��� Then
            mrs����ҩ������.Delete
        ElseIf Index = cbo_סԺҩ�����Ҷ��շ��� Then
            mrsסԺҩ������.Delete
        ElseIf Index = cbo_סԺ����ִ���Զ���ɷ��� Then
            mrsסԺִ�ж���.Delete
        End If
        cbo(Index).RemoveItem cbo(Index).ListIndex
        If cbo(Index).ListIndex = -1 Then
            If cbo(Index).ListCount > 0 Then
                cbo(Index).ListIndex = 0
            Else
                vsUseDept(Index).Rows = 0: vsUseDept(Index).Rows = 1
                vsUseDept(Index).Enabled = False
                If Index = cbo_סԺ����ִ���Զ���ɷ��� Then
                    lst(lst_����ִ���Զ����ҽ�����).Enabled = False
                Else
                    vsfDrugStore(Index).Enabled = False
                End If
            End If
        Else
            Call cbo_Click(Index)
        End If
        
        If Index = cbo_סԺ����ִ���Զ���ɷ��� Then
            Frame14.Tag = "���޸�"
        Else
            SST.Tag = "���޸�"
        End If
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim strSQL As String
    
    If vsfMecItem.Row > 0 Then
        If CheckMecItem = False Then vsfMecItem.SetFocus: Exit Sub
        If MsgBox("ȷ��Ҫɾ��[" & vsfMecItem.TextMatrix(vsfMecItem.Row, 1) & "]��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        strSQL = "zl_������Ŀ_edit(null,null,null,'" & vsfMecItem.TextMatrix(vsfMecItem.Row, 0) & "',2)"
        On Error GoTo ErrHandle
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        vsfMecItem.RemoveItem vsfMecItem.Row
        If vsfMecItem.Rows = 1 Then
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDepartSelect_Click()
    Dim lngCol As Long, lngRow As Long, X As Long, Y As Long, rsTmp As ADODB.Recordset, blnCancel As Boolean
    Dim i As Long, blnChange As Boolean, strNewValue As String
    
    lngCol = vsfDepartSign.MouseCol: lngRow = vsfDepartSign.MouseRow
    If lngRow < 1 Then Exit Sub
    vsfDepartSign.Row = vsfDepartSign.Rows - 1 'ѡ���޸���
    gstrSQL = "Select b.Id, b.����, b.����,b.���롡" & vbNewLine & _
              "from (Select ����id, ����id, ����ֵ From Zldeptparas Where ����id = (Select ID From zlParameters Where ������ = 'ǩ��ʹ��ͼƬ')) A,���ű� B" & vbNewLine & _
              "Where b.Id = a.����id(+) And a.����id Is Null" & vbNewLine & _
              "And (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
              "Order By b.����"

    Set rsTmp = FS.ShowSQLSelectEx(Me, cmdDepartSelect, gstrSQL, 0, "", False, "����", "��ѡ����", False, False, True, blnCancel, False, False, False, "")
    
    If blnCancel = False Then
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.EOF Then Exit Sub
        For i = 1 To vsfDepartSign.Rows - 1
            If vsfDepartSign.TextMatrix(i, vsfDepartSign.ColIndex("ID")) = rsTmp!ID Then Exit Sub
        Next
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("ID")) = rsTmp!ID & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("����")) = rsTmp!���� & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("����")) = rsTmp!���� & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("����")) = "1"
        blnChange = True
    End If
        
    If blnChange Then
        If vsfDepartSign.TextMatrix(vsfDepartSign.Rows - 1, vsfDepartSign.ColIndex("ID")) <> "" Then
            vsfDepartSign.Rows = vsfDepartSign.Rows + 1
        End If

        With vsfDepartSign
            For i = 1 To .Rows - 1
                If Val(.Cell(flexcpChecked, i, .ColIndex("����"))) <> Decode(Val(.RowData(i)), 1, 1, 2) Then
                    vsfDepartSign.ForeColor = vbRed
                    Exit Sub
                End If
            Next
            vsfDepartSign.ForeColor = vbBlack
        End With
        cmdDepartSelect.Visible = False
        vsfDepartSign.Row = vsfDepartSign.Rows - 1
        Call vsfDepartSign.ShowCell(vsfDepartSign.Rows - 1, 0)
    End If
End Sub

Private Sub cmdEprDown_Click()
    If vsfEpr(0).Row < vsfEpr(0).Rows - 1 Then
        Dim intCount As Integer, strTmp As String
        '����д洢ֵ
        strTmp = CStr(vsfEpr(0).RowData(vsfEpr(0).Row))
        vsfEpr(0).RowData(vsfEpr(0).Row) = vsfEpr(0).RowData(vsfEpr(0).Row + 1)
        vsfEpr(0).RowData(vsfEpr(0).Row + 1) = Val(strTmp)
        '�������ʾֵ
        strTmp = vsfEpr(0).TextMatrix(vsfEpr(0).Row, 0)
        vsfEpr(0).TextMatrix(vsfEpr(0).Row, 0) = vsfEpr(0).TextMatrix(vsfEpr(0).Row + 1, 0)
        vsfEpr(0).TextMatrix(vsfEpr(0).Row + 1, 0) = strTmp
        '�����ǰѡ����
        vsfEpr(0).Row = vsfEpr(0).Row + 1
        '���ֵ
        strTmp = ""
        With vsfEpr(0)
            For intCount = 1 To .Rows - 1
                If Val(.RowData(intCount)) > 0 Then
                    strTmp = strTmp & ";" & Val(.RowData(intCount))
                End If
            Next
        End With
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        Call SetParChange(vsfEpr, 0, mrsPar, True, strTmp)
    End If
End Sub

Private Sub cmdEprSelect_Click(Index As Integer)
Dim lngCol As Long, lngRow As Long, X As Long, Y As Long, rsTmp As New ADODB.Recordset, blnCancel As Boolean
Dim vPoint As POINTAPI, l As Long, blnChange As Boolean, strNewValue As String
    vPoint = zlControl.GetCoordPos(vsfEpr(Index).hwnd, cmdEprSelect(Index).Left, cmdEprSelect(Index).Top)
    lngCol = vsfEpr(Index).MouseCol: lngRow = vsfEpr(Index).MouseRow
    If lngRow < 1 Then Exit Sub
    If lngCol = 1 Then '����
        gstrSQL = "Select Distinct a.���,a.Id, a.����, a.����, c.���� As ����" & vbNewLine & _
                "From ��Ա�� A, ��Ա����˵�� B, ���ű� C, ������Ա D" & vbNewLine & _
                "Where a.Id = b.��Աid And c.Id = d.����id And d.��Աid = a.Id And d.ȱʡ = 1 And" & vbNewLine & _
                "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And b.��Ա���� In ('ҽ��')" & vbNewLine & _
                "Order By a.���"
        Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "", False, "����", "��ѡ��һ�������Ա", False, False, True, vPoint.X, vPoint.Y, , blnCancel)
        If blnCancel = False Then
            If rsTmp.EOF Then Exit Sub
            blnChange = True
            vsfEpr(Index).TextMatrix(lngRow, 0) = rsTmp!ID
            vsfEpr(Index).TextMatrix(lngRow, 1) = rsTmp!����
        End If
    ElseIf lngCol = 3 Then '����
        gstrSQL = "Select a.Id,a.����, a.����, a.����" & vbNewLine & _
                    "From ���ű� A, ��������˵�� B" & vbNewLine & _
                    "Where a.Id = b.����id And b.�������� In ('�ٴ�') And b.������� In (2, 3) And" & vbNewLine & _
                    "      (To_Char(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' Or a.����ʱ�� Is Null)" & vbNewLine & _
                    "Order By a.����"
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, gstrSQL, 0, "", False, "", "��ѡ��һ���������˿���", False, False, True, vPoint.X, vPoint.Y, 225, blnCancel, True, True)
        If blnCancel = False Then
            If rsTmp.EOF Then Exit Sub
            blnChange = True
            Do Until rsTmp.EOF
                If rsTmp.AbsolutePosition = 1 Then
                    vsfEpr(Index).TextMatrix(lngRow, 2) = rsTmp!ID
                    vsfEpr(Index).TextMatrix(lngRow, 3) = rsTmp!����
                Else
                    vsfEpr(Index).TextMatrix(lngRow, 2) = vsfEpr(Index).TextMatrix(lngRow, 2) & "," & rsTmp!ID
                    vsfEpr(Index).TextMatrix(lngRow, 3) = vsfEpr(Index).TextMatrix(lngRow, 3) & vbCrLf & rsTmp!����
                End If
                rsTmp.MoveNext
            Loop
            vsfEpr(Index).AutoSize 3, 3
        End If
    End If
    
    If blnChange Then
        If vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 0) <> "" And vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 2) <> "" Then
            vsfEpr(Index).Rows = vsfEpr(Index).Rows + 1
        End If
        
        For l = 1 To vsfEpr(Index).Rows - 1
            If vsfEpr(Index).TextMatrix(l, 0) <> "" And vsfEpr(Index).TextMatrix(l, 2) <> "" Then
                strNewValue = strNewValue & ";" & vsfEpr(Index).TextMatrix(l, 0) & "," & vsfEpr(Index).TextMatrix(l, 2)
             End If
        Next
        If strNewValue <> "" Then
            strNewValue = Mid(strNewValue, 2)
        End If
        
        Call SetParChange(vsfEpr, Index, mrsPar, True, strNewValue)
    End If
End Sub


Private Sub cmdEprUp_Click()
    If vsfEpr(0).Row > 1 Then
        Dim intCount As Integer, strTmp As String
        '����д洢ֵ
        strTmp = CStr(vsfEpr(0).RowData(vsfEpr(0).Row))
        vsfEpr(0).RowData(vsfEpr(0).Row) = vsfEpr(0).RowData(vsfEpr(0).Row - 1)
        vsfEpr(0).RowData(vsfEpr(0).Row - 1) = Val(strTmp)
        '�������ʾֵ
        strTmp = vsfEpr(0).TextMatrix(vsfEpr(0).Row, 0)
        vsfEpr(0).TextMatrix(vsfEpr(0).Row, 0) = vsfEpr(0).TextMatrix(vsfEpr(0).Row - 1, 0)
        vsfEpr(0).TextMatrix(vsfEpr(0).Row - 1, 0) = strTmp
        '�����ǰѡ����
        vsfEpr(0).Row = vsfEpr(0).Row - 1
        '���ֵ
        strTmp = ""
        With vsfEpr(0)
            For intCount = 1 To .Rows - 1
                If Val(.RowData(intCount)) > 0 Then
                    strTmp = strTmp & ";" & Val(.RowData(intCount))
                End If
            Next
        End With
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        Call SetParChange(vsfEpr, 0, mrsPar, True, strTmp)
    End If
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdLink_Click()
    Call CheckNurseIntegrateIP
End Sub

Private Function CheckNurseIntegrateIP(Optional ByVal blnComdValiade As Boolean = True) As Boolean
    Dim strIP As String, strErrMsg As String
    If txt(txt_���廤��IP��ַ).Text = "" Then
        If blnComdValiade = True Then
            MsgBox "���������廤��IP��ַ!", vbInformation, gstrSysName
            Call zlControl.ControlSetFocus(txt(txt_���廤��IP��ַ))
            Exit Function
        End If
    End If
    '�ж�IP��ַ�Ƿ���ȷ
    If IsIPAddress(txt(txt_���廤��IP��ַ).Text) = False Then
        If blnComdValiade = True Then
            MsgBox "IP��ַ��ʽ����ȷ�����������룡", vbInformation, gstrSysName
            Call zlControl.ControlSetFocus(txt(txt_���廤��IP��ַ))
        End If
        Exit Function
    End If
    '����Ƿ�����ȷ�ƶ���������ַ�����ӷ��������в���
    strIP = txt(txt_���廤��IP��ַ).Text & IIF(txt(txt_���廤��IP�˿�).Text = "", "", ":" & txt(txt_���廤��IP�˿�).Text)
    If InitNurseIntegrate(True) = True Then
        If gobjNurseIntegrate.IPAdreesCheck(strIP, strErrMsg) = False Then
            If blnComdValiade = True Then
                MsgBox strErrMsg, vbInformation, gstrSysName
            End If
            Exit Function
        Else
            If blnComdValiade = True Then MsgBox "IP��ַ���óɹ���", vbInformation, gstrSysName
            cmdLink.Tag = "OK"
        End If
        CheckNurseIntegrateIP = True
    End If
End Function

Private Sub cmdModify_Click()
    If vsfMecItem.Row > 0 Then
        If CheckMecItem = False Then vsfMecItem.SetFocus: Exit Sub
        frmInMedSetup.ShowMe vsfMecItem.TextMatrix(vsfMecItem.Row, 0), vsfMecItem.TextMatrix(vsfMecItem.Row, 1), vsfMecItem.TextMatrix(vsfMecItem.Row, 2), "�޸�", Me
        vsfMecItem.SetFocus
    End If
End Sub

Private Function CheckMecItem() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select count(��Ϣ��) as ���� from ������ҳ�ӱ� where ��Ϣ��='" & vsfMecItem.TextMatrix(vsfMecItem.Row, 1) & "'"
    
    Err = 0: On Error GoTo ErrHandle
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
    If rsTemp!���� > 0 Then
        MsgBox "����Ŀ�Ѿ�ʹ��,���ܽ����޸Ļ�ɾ��!"
        CheckMecItem = False
        Exit Function
    End If
    CheckMecItem = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdPathSortSet_Click()
    frmParAdviceSort.mbytFun = 0
    frmParAdviceSort.Show vbModal, Me
End Sub


Private Sub cmdPatiSign_Click()
    frmPatiStamp.ShowMe Me
End Sub

Private Sub cmdSet_Click()
    Call mobjPass.Setup(Me, gcnOracle, glngSys)
End Sub

Private Sub cmdUnVisible_Click(Index As Integer)
    Dim lngIndex As Long
    
    picControl(Index).Visible = False
    If picControl(Index).Tag <> "" Then
        lngIndex = Val(picControl(Index).Tag)
        If picLineColor(lngIndex).Enabled And picLineColor(lngIndex).Visible Then picLineColor(lngIndex).SetFocus
    End If
    picControl(Index).Tag = ""
End Sub

Private Sub cmdҩƷ�����ҩ;��_Click(Index As Integer)
    Call SetLstSelected(lst(lst_ҩƷ�����ҩ;��), Index = 0)
End Sub

Private Sub cmd����ִ���Զ����ҽ�����_Click(Index As Integer)
    If lst(lst_����ִ���Զ����ҽ�����).Enabled = False Then Exit Sub
    Call SetLstSelected(lst(lst_����ִ���Զ����ҽ�����), Index = 0)
End Sub


Private Sub Form_Activate()
    If Me.Tag = "��ʼ�ɹ�" Then
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    
    mblnOk = False
    strCategory = "��������,������Ŀ"
    
    'ͼ����,TaskPanelItem��ID(ͬʱҲ�ǲ�������Picture�ؼ������),TaskPanelItem�ı���;......
    marrFunc(0) = "100,0,ҽ���´�ѡ��;101,1,ҵ�����̿���;102,2,����ҽ������;103,3,סԺҽ������;104,4,ҽ����������;" & _
                  "106,6,�ٴ�·������;107,8,�ٴ�����վ;108,7,סԺ��ҳ;105,9,������д��ʾ;107,10,���Ӳ�������;" & _
                  "103,11,������Һ����;112,12,�°滤���ļ�"
    marrFunc(1) = "105,5,����ҩ������"

    '1.��ʼ���������һ�������б�,ȱʡѡ�е�һ��
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.��ʼ���������Ķ��������б�,ȱʡѡ�е�һ��
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    Me.Tag = "��ʼ�ɹ�"
End Sub


Private Sub OptEnemaStool_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(OptEnemaStool, Index, mrsPar)
End Sub

Private Sub OptEnemaStool_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub OptEnemaStool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(OptEnemaStool, Index, mrsPar)
End Sub

Private Sub optEprRead_Click(Index As Integer)
    If Index = 2 Then
        txt(txt_����������Ԥ��N��).Enabled = True
        Call SetParChange(txt, txt_����������Ԥ��N��, mrsPar, True, Val(txt(txt_����������Ԥ��N��)))
        Call SetParChange(optEprRead, Index, mrsPar, True, Val(txt(txt_����������Ԥ��N��)))
    Else
        txt(txt_����������Ԥ��N��).Enabled = False
        Call SetParChange(optEprRead, Index, mrsPar, True, IIF(Index = 0, "-1", "0"))
    End If
End Sub

Private Sub optEprRead_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optEprRead, Index, mrsPar)
End Sub

Private Sub optFileTime_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optFileTime, Index, mrsPar)
End Sub

Private Sub optFileTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optFileTime_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optFileTime, Index, mrsPar)
End Sub

Private Sub optICD����_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optICD����, Index, mrsPar)
End Sub

Private Sub optICD����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optICD����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optICD����, Index, mrsPar)
End Sub

Private Sub OptInsert_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(OptInsert, Index, mrsPar)
End Sub

Private Sub OptInsert_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub OptInsert_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(OptInsert, Index, mrsPar)
End Sub

Private Sub optNewCard_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optNewCard, Index, mrsPar)
End Sub

Private Sub optNewCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optNewCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optNewCard, Index, mrsPar)
End Sub

Private Sub OptOut_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(OptOut, Index, mrsPar)
End Sub

Private Sub OptOut_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub OptOut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(OptOut, Index, mrsPar)
End Sub

Private Sub optPloy_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPloy, Index, mrsPar)
End Sub

Private Sub optPloy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPloy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPloy, Index, mrsPar)
End Sub

Private Sub optSign_Click(Index As Integer)
    Call SetParChange(optSign, Index, mrsPar)
End Sub

Private Sub optSign_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSign, Index, mrsPar)
End Sub

Private Sub opt���˹���_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt���˹���, Index, mrsPar)
End Sub

Private Sub opt���˹���_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt���˹���, Index, mrsPar)
End Sub

Private Sub opt·����ӡ����_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt·����ӡ����, Index, mrsPar)
End Sub

Private Sub opt·����ӡ����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt·����ӡ����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt·����ӡ����, Index, mrsPar)
End Sub

Private Sub optPrintWay_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPrintWay, Index, mrsPar)
End Sub

Private Sub optPrintWay_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintWay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintWay, Index, mrsPar)
End Sub

Private Sub optÿҳ·����ӡ����_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optÿҳ·����ӡ����, Index, mrsPar)
End Sub

Private Sub optÿҳ·����ӡ����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optÿҳ·����ӡ����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optÿҳ·����ӡ����, Index, mrsPar)
End Sub

Private Sub opt����·����ӡ����_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt����·����ӡ����, Index, mrsPar)
End Sub

Private Sub opt����·����ӡ����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt����·����ӡ����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt����·����ӡ����, Index, mrsPar)
End Sub

Private Sub opt����·����ӡ����_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt����·����ӡ����, Index, mrsPar)
End Sub

Private Sub opt����·����ӡ����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt����·����ӡ����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt����·����ӡ����, Index, mrsPar)
End Sub

Private Sub opt����_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt����, Index, mrsPar)
End Sub

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt����, Index, mrsPar)
End Sub

Private Sub opt�����ж�_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt�����ж�, Index, mrsPar)
End Sub

Private Sub opt�����ж�_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt�����ж�_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt�����ж�, Index, mrsPar)
End Sub


Private Sub opt�������_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt�������, Index, mrsPar)
End Sub

Private Sub opt�������_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt�������_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt�������, Index, mrsPar)
End Sub

Private Sub optת��������Ժҽ��_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optת��������Ժҽ��, Index, mrsPar)
End Sub

Private Sub optת��������Ժҽ��_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optת��������Ժҽ��, Index, mrsPar)
End Sub

Private Sub optPrintDruUse_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPrintDruUse, Index, mrsPar)
End Sub

Private Sub optPrintDruUse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintDruUse, Index, mrsPar)
End Sub

Private Sub opt��ҩ����_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt��ҩ����, Index, mrsPar)
End Sub

Private Sub opt��ҩ����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt��ҩ����, Index, mrsPar)
End Sub

Private Sub opt�������_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt�������, Index, mrsPar, True, IIF(opt�������(0).value, "0|0", IIF(opt�������(1).value, "1|" & NVL(txt(txt_������ǰ�������).Text, "0"), "2|" & NVL(txt(txt_������ǰ�������).Text, "0"))))
    txt(txt_������ǰ�������).Enabled = Index > 0
End Sub

Private Sub opt�������_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt�������_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt�������, Index, mrsPar)
End Sub

Private Sub PicColorCollect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    
    If X > 0 And X < PicColorCollect(Index).ScaleWidth And Y > 0 And Y < PicColorCollect(Index).ScaleHeight Then
        SetCapture PicColorCollect(Index).hwnd
        shpBorder(Index).Visible = True
    Else
        ReleaseCapture
        shpBorder(Index).Visible = False
    End If

    
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    shpBorder(Index).Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    
    If PicColorCollect(Index).Point(lX, lY) = -1 Then Exit Sub
    picColor(Index).BackColor = PicColorCollect(Index).Point(lX, lY)
    Select Case CStr(Hex(picColor(Index).BackColor))
    Case "0"
        lblColor(Index) = "��ɫ"
    Case "3399"
        lblColor(Index) = "��ɫ"
    Case "3333"
        lblColor(Index) = "���ɫ"
    Case "3300"
        lblColor(Index) = "����"
    Case "663300"
        lblColor(Index) = "����"
    Case "800000"
        lblColor(Index) = "����"
    Case "993333"
        lblColor(Index) = "����"
    Case "333333"
        lblColor(Index) = "��ɫ-80%"
    Case "80"
        lblColor(Index) = "���"
    Case "66FF"
        lblColor(Index) = "��ɫ"
    Case "8080"
        lblColor(Index) = "���"
    Case "8000"
        lblColor(Index) = "��ɫ"
    Case "808000"
        lblColor(Index) = "��ɫ"
    Case "FF0000"
        lblColor(Index) = "��ɫ"
    Case "996666"
        lblColor(Index) = "��-��"
    Case "808080"
        lblColor(Index) = "��ɫ-50%"
    Case "FF"
        lblColor(Index) = "��ɫ"
    Case "99FF"
        lblColor(Index) = "ǳ��ɫ"
    Case "CC99"
        lblColor(Index) = "���ɫ"
    Case "669933"
        lblColor(Index) = "����"
    Case "CCCC33"
        lblColor(Index) = "ˮ��ɫ"
    Case "FF6633"
        lblColor(Index) = "ǳ��"
    Case "800080"
        lblColor(Index) = "������"
    Case "999999"
        lblColor(Index) = "��ɫ-40%"
    Case "FF00FF"
        lblColor(Index) = "�ۺ�"
    Case "CCFF"
        lblColor(Index) = "��ɫ"
    Case "FFFF"
        lblColor(Index) = "��ɫ"
    Case "FF00"
        lblColor(Index) = "����"
    Case "FFFF00"
        lblColor(Index) = "����"
    Case "FFCC00"
        lblColor(Index) = "����"
    Case "663399"
        lblColor(Index) = "÷��"
    Case "C0C0C0"
        lblColor(Index) = "��ɫ-25%"
    Case "CC99FF"
        lblColor(Index) = "õ���"
    Case "99CCFF"
        lblColor(Index) = "��ɫ"
    Case "99FFFF"
        lblColor(Index) = "ǳ��"
    Case "CCFFCC"
        lblColor(Index) = "ǳ��"
    Case "FFFFCC"
        lblColor(Index) = "ǳ����"
    Case "FFCC99"
        lblColor(Index) = "����"
    Case "FF99CC"
        lblColor(Index) = "����"
    Case "FFFFFF"
        lblColor(Index) = "��ɫ"
    Case Else
        lblColor(Index) = "&H" & CStr(Hex(picColor(Index).BackColor))
    End Select
End Sub

Private Sub PicColorCollect_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strValue As String
    Dim lngIndex As Long
    
    If Button <> vbLeftButton Then Exit Sub
    '��ָ����ɫ��ͼ
    If picControl(Index).Tag = "" Then
        picControl(Index).Visible = False
        Exit Sub
    End If
    lngIndex = Val(picControl(Index).Tag)
    picLineColor(lngIndex).BackColor = picColor(Index).BackColor
    If Me.Visible Then
        Call SetParChange(picLineColor, lngIndex, mrsPar, True, Val(picLineColor(lngIndex).BackColor))
    End If
    picControl(Index).Visible = False
     picControl(Index).Tag = ""
    If picLineColor(lngIndex).Enabled And picLineColor(lngIndex).Visible Then picLineColor(lngIndex).SetFocus
End Sub

Private Sub picLineColor_Click(Index As Integer)
    Dim lngIndex As Long
    
    lngIndex = Index
    If lngIndex = 1 Or lngIndex = 2 Or lngIndex = 3 Or lngIndex = 4 Then lngIndex = 1 'ȷ����ɫѡ��ؼ�����
    
    picControl(lngIndex).Top = picLineColor(Index).Top + picLineColor(Index).Height
    picControl(lngIndex).Left = picLineColor(Index).Left
    If picControl(lngIndex).Left + picControl(lngIndex).Width > StabNurse.Width - 60 Then
        picControl(lngIndex).Left = StabNurse.Width - 60 - picControl(lngIndex).Width
    End If
    If picControl(lngIndex).Left < 60 Then picControl(lngIndex).Left = 60
    picControl(lngIndex).Visible = True
    picControl(lngIndex).ZOrder 0
    picControl(lngIndex).Tag = Index
    Call SetCOLOR(Val(picLineColor(Index).BackColor), lngIndex)
End Sub

Private Sub picLineColor_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub picLineColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(picLineColor, Index, mrsPar)
End Sub

Private Sub picPar_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 11     '������Һ����
        With vsfWaittingMixDept
            .Width = picPar(Index).ScaleWidth - 360 * 2
            .Height = picPar(Index).ScaleHeight - .Top - 360
        End With
    End Select
End Sub

Private Sub StabNurse_Click(PreviousTab As Integer)
    If Me.Visible Then
        If StabNurse.Tab = 1 Then
            lblPrompt.Caption = "˵��:�˽���ֻ��Ա�׼ͨ�����²���,������������ģ���ڲ��Ļ���ѡ��������"
        Else
            lblPrompt.Caption = ""
        End If
    End If
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim i As Long
    
    For i = 0 To picPar.UBound
        picPar(i).Visible = (i = Item.ID)
    Next
    
    lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("ҵ�����̿���", marrFunc) Or Item.ID = GetFuncID("סԺҽ������", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
        If Item.ID = GetFuncID("ҵ�����̿���", marrFunc) Then
            lblLocate(txt_Dept).Caption = "���Ҳ���(&F)"
        Else
            lblLocate(txt_Dept).Caption = "��������(&F)"
        End If
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOK.Left - lblPrompt.Left - 120
    If Item.ID = GetFuncID("�°滤���ļ�", marrFunc) Then
        Call StabNurse_Click(StabNurse.Tab)
    Else
        lblPrompt.Caption = ""
    End If
    mlngPreFind = 1
    
    tplFunc.Tag = Item.ID   '���ڻ�ȡ��ǰѡ�е�TaskPanelItem
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picVbar.Left < 1500 Then picVbar.Left = 1500
    If picVbar.Left > Me.ScaleWidth - 3000 Then picVbar.Left = Me.ScaleWidth - 3000
    picVbar.Top = 0
    
    picFunc.Width = picVbar.Left + picVbar.Width
    
    For i = 0 To picPar.UBound
        picPar(i).Top = Me.ScaleTop
        picPar(i).Left = picFunc.Left + picFunc.ScaleWidth
        picPar(i).Width = Me.ScaleWidth - picPar(i).Left
        picPar(i).Height = Me.ScaleHeight - PicBottom.ScaleHeight
    Next
End Sub


Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = PicBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
End Sub


Private Sub picFunc_Resize()
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    scbFunc.Height = picFunc.ScaleHeight
    
    picVbar.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub


Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIF(picVbar.Left + X < 2000, 2000, picVbar.Left + X)
        Call Form_Resize
    End If
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID�Ǵ�1��ʼ�ģ���ΪͬʱΪͼ����ţ�,�����Ǵ�0��ʼ
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub


Public Sub LocateFuncItem(ByVal lngFunc As Long)
'���ܣ�����IDѡ��һ���Ͷ�������
    Dim i As Long, j As Long, lngId As Long
    Dim arrTmp As Variant
    Dim n As Long
    
    For i = 0 To UBound(marrFunc)
        arrTmp = Split(marrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            lngId = Split(arrTmp(j), ",")(1)
            If lngFunc = lngId Then
                tplFunc.Tag = lngId
                Set scbFunc.Selected = scbFunc(i)
                
                For n = 1 To tplFunc.Groups(1).Items.Count
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).ID = lngId
                Next
            End If
        Next
    Next
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOk Then
        mrsPar.Filter = "(�޸�״̬=1 ANd ErrType =Null) OR  (�޸�״̬=1 And ErrType=" & PET_ֵ���� & ")"
        If mrsPar.RecordCount > 0 Or cmdAdvice.Tag = "���޸�" Or SST.Tag = "���޸�" Or Frame14.Tag = "���޸�" Then
            If MsgBox("�����޸Ĳ��ֲ����������������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mrsAdvice = Nothing
    Set mcol���� = Nothing
    Set mcolStop���� = Nothing
    Set mrsPar = Nothing
    Set mrs����ҩ������ = Nothing
    Set mrsסԺҩ������ = Nothing
    Set mrsסԺִ�ж��� = Nothing
    Set mobjPass = Nothing
    Set gobjNurseIntegrate = Nothing
End Sub

Private Sub InitData()
'���ܣ���ʼ������ؼ�,��ȡ����������
    
    '1.��ʼ������
    mlngPreFind = 1
    mblnUseBlood = isPiotBlood
    Set mcol���� = New Collection
    Set mcolStop���� = New Collection
    
    Call InitSystemPara
    
    
    '2.��ʼ������ؼ�
    Call InitEnv
    
    
    '3.����ϵͳ����
    Call LoadPar
    
    '4.�������ݺ�Ŀؼ������Կ���
    chk_Click chk_סԺ�´��Զ�����
    chk_Click chk_����ǰһ�첻���������ɽ���·����Ŀ
    chk_Click chk_����·��ִ�л���
    chk_Click chk_������ת�ƻ�ҳ
    chk_Click chk_סԺ�����Զ�ִ�г���
    chk_Click chk_����ҩ�����հ�������������
End Sub
Private Sub LoadPar()
'���ܣ���ȡ�����ز���������ؼ�
    Dim strValue As String, strTmp As String, arrTmp As Variant, varTmp As Variant
    Dim i As Long, n As Long, intIndex As Integer, lngValue As Long
    Dim rsTmp As ADODB.Recordset, rs As ADODB.Recordset, rsData As ADODB.Recordset
    Dim arrObj As Variant  '�������ģ��1,������1(������1),�ؼ�����1,ģ��2,������2(������2),�ؼ�����2,......,ģ�������ʹ�ò�����
    
    Set rsTmp = GetPar(mrsPar, p����ҽ���´� & "," & pסԺҽ���´� & "," & pסԺҽ������ & "," & p�ٴ�·��Ӧ�� & _
        "," & pסԺҽ��վ & "," & p����ҽ��վ & "," & pסԺ��ʿվ & "," & p�°�סԺ��ʿվ & "," & pҽ������վ & "," & p�����ڲ����� & _
        "," & pסԺ�������� & "," & p���Ӳ������ & "," & p���Ӳ������� & "," & p���Ӳ������� & "," & p������Һ���� & _
        "," & p�����¼���� & "," & p�ٴ�·������ & "," & p����·��Ӧ��)
    
     '1.����CheckBox�����
    strTmp = "0:27:" & chk_סԺҩ�����Ͳ�����ҩ�� & _
            ",0:34:" & chk_ָ��ҽ������������ִ�� & _
            ",0:43:" & chk_�´��Ժҽ���������Ժ & _
            ",0:51:" & chk_����ִ�еǼ� & _
            ",0:56:" & chk_���ﴦ���������� & _
            ",0:68:" & chk_δ����������ֹ��ҩ & _
            ",0:69:" & chk_ҩƷ�������ҽ�� & _
            ",0:70:" & chk_�����Ǽ���Ч���� & _
            ",0:71:" & chk_����ҽ��������Ч & _
            ",0:84:" & chk_һ��������������Ŀ & _
            ",0:143:" & chk_����ҽ���������������� & _
            ",0:161:" & chk_����ҩ�� & _
            ",0:162:" & chk_�´�ҽ��ʱ��ʾ���� & _
            ",0:182:" & chk_��ֹ�´ﳬ����ҩƷҽ�� & _
            ",0:187:" & chk_����ҩ��ּ����� & _
            ",0:188:" & chk_����ҩ��ʹ���Ա�ҩ & _
            ",0:189:" & chk_����Ժ��ִ�н���ҩƷ & _
            ",0:192:" & chk_���˳�Ժҽ������������Ժ & _
            ",0:271:" & chk_ͣ��ʱ¼��ԭ�� & _
            ",0:274:" & chk_����ҩ�����հ������������� & _
            ",0:288:" & chk_������Ҫ����ִ�� & _
            ",0:300:" & chk_��Ⱦ�����濨ǿ����д & _
            ",0:302:" & chk_��������´�ҽ���ɻ���������Ҵ���

    strTmp = strTmp & _
            ",0:208:" & chk_�ٴ�����վ����ʹ��zlPlugIn���� & _
            ",0:209:" & chk_�����ּ����� & _
            ",0:210:" & chk_���������Һ���Ч�����Ĳ��� & _
            ",0:216:" & chk_��Ѫ�ּ����� & _
            ",0:217:" & chk_������Ȩ���� & _
            ",0:218:" & chk_��Ѫ����������� & _
            ",0:219:" & chk_��Ѫ���������м�������ҽʦ & _
            ",0:225:" & chk_�ӿڵ�����־_��ͨ & _
            ",0:226:" & chk_ʹ��ϵͳ����_���� & _
            ",0:230:" & chk_ҽ������ʱ��������ԭ�� & _
            ",0:236:" & chk_����Ѫ�����ϵͳ & _
            ",0:237:" & chk_��ƻ��������дҪ�� & _
            ",0:240:" & chk_ҽ����ҩ�������� & _
            ",0:247:" & chk_���֤���� & _
            ",0:248:" & chk_��ҽ��С����п���ҩ����� & _
            ",0:249:" & chk_����ҩƷҪ����дԭ�� & _
            ",0:250:" & chk_�����ּ���� & _
            ",0:254:" & chk_����ҽʦ�ﵽ�����ȼ�������� & _
            ",0:257:" & chk_��������������ɵ��� & _
            ",0:259:" & chk_·������ԭ����ֵ����ѡȡ & _
            ",0:262:" & chk_����ҩƷ�ֿ����� & _
            ",0:272:" & chk_��Ѫҽ�����ͺ���ܷ�Ѫ & _
            ",0:286:" & chk_��Ѫ���벻��ʾѪҺ��� & _
            ",0:289:" & chk_���¼��ʱ�����Զ���ȡ
    strTmp = strTmp & _
            ",0:293:" & chk_�´���Ѫ����ʱȷ����Ѫ��Ϣ & _
            ",0:301:" & chk_ѪҺ���պ������ִ�еǼ� & _
            ",0:307:" & chk_����д�����

    strTmp = strTmp & _
            "," & p����ҽ���´� & ":ҽ��ִ������:" & chk_����ҩ�������� & _
            "," & p����ҽ���´� & ":�Զ�����Ƥ��ҽ��:" & chk_�����´��Ƥ�� & _
            "," & p����ҽ���´� & ":����¼��ҩƷ����:" & chk_����ҩ���䵥�� & _
            "," & p����ҽ���´� & ":��λ����:" & chk_��Լ��λ���ͼ��� & _
            "," & p����ҽ���´� & ":���ﱾ���Զ�ִ��:" & chk_���﷢�ͱ����Զ�ִ�� & _
            "," & p����ҽ���´� & ":Ҫ��ǼǴ�����:" & chk_����ǼǴ����� & _
            "," & p����ҽ���´� & ":һ����ҩ����Ϊһ��:" & chk_һ������һ�ŵ��� & _
            "," & p����ҽ���´� & ":��ͬ��ϵ�ҽ���ֱ��������:" & chk_��ͬ��Ϸֵ��� & _
            "," & p����ҽ���´� & ":����ҽ�����ͺ��������֧��:" & chk_����ҽ�����ͺ��������֧�� & _
            "," & p����ҽ���´� & ":��ʼʱ�䲻��ͬһ��ķֱ��������:" & chk_��ͬ��ķֵ��� & _
            "," & p����ҽ���´� & ":��ʾҩƷ���:" & chk_����ҽ���´�ʱ����ѡ������ʾҩƷ��� & _
            "," & p����ҽ���´� & ":����ҽ��������������:" & chk_�������ҽ������ʱһ����鷢��Ϊһ�ŵ��� & _
            "," & p����ҽ���´� & ":ҽ������Ƥ������:" & chk_����ҽ������Ƥ������ & _
            "," & p����ҽ���´� & ":Ƥ��������ҩ:" & chk_Ƥ��������ҩ���� & _
            "," & p����ҽ���´� & ":������ҽ������¼����ҽ���:" & chk_������ҽ������¼����ҽ���
    
    strTmp = strTmp & _
            "," & pסԺҽ���´� & ":����ȱʡһ����:" & chk_סԺ����ȱʡһ���� & _
            "," & pסԺҽ���´� & ":���������뵥��:" & chk_סԺҩ���䵥�� & _
            "," & pסԺҽ���´� & ":ҽ��ִ������:" & chk_סԺҩ�������� & _
            "," & pסԺҽ���´� & ":Ҫ�������Ժ���:" & chk_�³�Ժҽ������Ժ��� & _
            "," & pסԺҽ���´� & ":������ɺ��´�����ҽ��:" & chk_����ִ�к�������ҽ�� & _
            "," & pסԺҽ���´� & ":ҽ���Զ�����:" & chk_סԺ�´��Զ����� & _
            "," & pסԺҽ���´� & ":ʵϰҽ��ֹͣҽ����Ҫ���:" & chk_ʵϰҽ��ͣ������� & _
            "," & pסԺҽ���´� & ":���������ס�����´�ҽ��:" & chk_����ס����ҽ���´� & _
            "," & pסԺҽ���´� & ":�Զ�����Ƥ��ҽ��:" & chk_סԺ�´��Ƥ�� & _
            "," & pסԺҽ���´� & ":��ʾҩƷ���:" & chk_סԺҽ���´�ʱ����ѡ������ʾҩƷ��� & _
            "," & pסԺҽ���´� & ":������ɺ�ر�ҽ������:" & chk_������ɺ�ر�ҽ������ & _
            "," & pסԺҽ���´� & ":Ҫ��ǼǴ�����:" & chk_סԺ�ǼǴ����� & _
            "," & pסԺҽ���´� & ":ҽ������Ƥ������:" & chk_סԺҽ������Ƥ������ & _
            "," & pסԺҽ���´� & ":Ƥ��������ҩ:" & chk_Ƥ��������ҩסԺ
            
   strTmp = strTmp & _
            "," & pסԺҽ������ & ":�Զ�����ҽ����ӡ:" & chk_У��ȷ��ֹͣ�������ӡ & _
            "," & pסԺҽ������ & ":����ҽ��У��:" & chk_����У�� & _
            "," & pסԺҽ������ & ":����ҽ����ͣ:" & chk_������ͣ & _
            "," & pסԺҽ������ & ":Ƥ����֤���:" & chk_�Ǽ�Ƥ����֤��� & _
            "," & pסԺҽ������ & ":ҽ��ҽ����������:" & chk_������ҽ��ҽ�� & _
            "," & pסԺҽ������ & ":�����ջط��ñ����Զ����:" & chk_�����ջ��Զ���˱��� & _
            "," & pסԺҽ������ & ":У��ҽ������ǩ��:" & chk_У��ȷ��ֹͣ����ǩ�� & _
            "," & pסԺҽ������ & ":ҩ���������ƽ���ʱ��:" & chk_ҩ���������ƽ���ʱ�� & _
            "," & pסԺҽ������ & ":���ҽ������:" & chk_סԺ���ͼ��ҽ������ & _
            "," & pסԺҽ������ & ":����ǰ�Զ�У��:" & chk_סԺ����У�Է��� & _
            "," & pסԺҽ������ & ":ֹͣ���Զ������ջ�:" & chk_ȷ��ֹͣ���Զ��ջ� & _
            "," & pסԺҽ������ & ":����ҽ������ǰ���δ��Чҽ��:" & chk_����ҽ�����ͼ��δ��Чҽ�� & _
            "," & pסԺҽ������ & ":���ñ��������ӡ:" & chk_ִ�е���ӡʱ������˻�ҳ��ӡ & _
            "," & pסԺҽ������ & ":������ת�ƻ�ҳ:" & chk_������ת�ƻ�ҳ & _
            "," & pסԺҽ������ & ":����������ҳ:" & chk_����������ҳ & _
            "," & pסԺҽ������ & ":������������ҳ:" & chk_������������ҳ & _
            "," & pסԺҽ������ & ":ת�ƻ�ҳ�������д�ӡ�ؿ�ҽ��:" & chk_ת�ƻ�ҳ�������д�ӡ�ؿ�ҽ�� & _
            "," & pסԺҽ������ & ":ֻ��ʾ��ǰ������ҽ��:" & chk_�ջ�ʱҽ���б�ֻ��ʾ��ǰ������ҽ�� & _
            "," & pסԺҽ������ & ":����ҽ��������������:" & chk_סԺ����ҽ������ʱһ����鷢��Ϊһ�ŵ��� & _
            "," & pסԺҽ������ & ":������ת�ƻ�ҳ:" & chk_������ת�ƻ�ҳ & _
            "," & pסԺҽ������ & ":����������ҳ:" & chk_����������ҳ & _
            "," & pסԺҽ������ & ":����δ����ҽ��ʱ��ֹ����ת��ҽ��:" & chk_����δ����ҽ��ʱ��ֹ����ת��ҽ��
    
    strTmp = strTmp & _
            "," & p�ٴ�·��Ӧ�� & ":�Ƿ�����·��ִ�л���:" & chk_����·��ִ�л��� & _
            "," & p�ٴ�·��Ӧ�� & ":ҽ��ҽ����·������:" & chk_·������ҽ��ҽ����ʾ & _
            "," & p�ٴ�·��Ӧ�� & ":δ����ʱ�������ҽ��������:" & chk_δ����ʱ�������ҽ�������� & _
            "," & p�ٴ�·��Ӧ�� & ":����ǰһ�첻���������ɽ����·����Ŀ:" & chk_����ǰһ�첻���������ɽ���·����Ŀ & _
            "," & p�ٴ�·��Ӧ�� & ":������ǰ���������·����Ŀ:" & chk_������ǰ���������·����Ŀ & _
            "," & p�ٴ�·��Ӧ�� & ":ƥ��ʱ��Ч��ͬ��·������Ŀ:" & chk_ƥ��ʱ��Ч��ͬ��·������Ŀ & _
            "," & p�ٴ�·��Ӧ�� & ":��Ժ������ȡ�����·��:" & chk_��Ժ������ȡ�����·�� & _
            "," & p�ٴ�·��Ӧ�� & ":ҩƷҽ����ƥ��Ϊ·������Ŀ:" & chk_ҩƷҽ����ƥ��Ϊ·������Ŀ & _
            "," & p�ٴ�·��Ӧ�� & ":ҩƷҽ����ͬ���಻��·����ҽ��:" & chk_ҩƷҽ����ͬ���಻��·����ҽ��
            
    strTmp = strTmp & "," & p�ٴ�·������ & ":˫���ģʽ:" & chk_����ҩ���ƺ�ҽ���˫���ģʽ
     
    strTmp = strTmp & _
            "," & p����·��Ӧ�� & ":�Ƿ�����·��ִ�л���:" & chk_��������·��ִ�л���
    strTmp = strTmp & _
            "," & pסԺҽ��վ & ":ʹ����������ʱ��:" & chk_ʹ����������ʱ�� & _
            "," & pסԺҽ��վ & ":��ҽ���Ҳ�ʹ����ҽ������ҳ��Ŀ:" & chk_��ҽ���Ҳ�ʹ����Ŀ & _
            "," & pסԺҽ��վ & ":ҽ���ͻ�ʿ�ֱ���д������ҳ:" & chk_ҽ���ͻ�ʿ�ֱ���д������ҳ & _
            "," & pסԺҽ��վ & ":����ʾ�޴�λ�Ĳ�������:" & chk_����ʾ�޴�λ�Ĳ������� & _
            "," & pסԺҽ��վ & ":�������ֻ����¼��������̬ѧ����:" & chk_�������ֻ����¼��M��ͷ��������̬ѧ���� & _
            "," & pסԺҽ��վ & ":������ϲ���Ϊ�����ٴ�·�����������:" & chk_������ϲ���Ϊ�����ٴ�·�����������
            
    strTmp = strTmp & _
            "," & p����ҽ��վ & ":��������������:" & chk_ҽ�������������ƺ����ﲡ�� & _
            "," & p����ҽ��վ & ":ҽ���������к���������:" & chk_ҽ���������к�������ڶ����н��� & _
            "," & p����ҽ��վ & ":ֻ�����Ѿ�����Ĳ���:" & chk_ֻ�����Ѿ�����Ĳ���
            
    strTmp = strTmp & _
            "," & pҽ������վ & ":δ�շ����:" & chk_����δ�շѲ������ִ�� & _
            "," & pҽ������վ & ":Ƥ����֤���:" & chk_��дƤ�Խ��ʱ��֤��� & _
            "," & pҽ������վ & ":ִ�б���ʱ�շѻ�������:" & chk_ִ�б���ʱ�շѻ������� & _
            "," & pҽ������վ & ":Ѫ͸����д�°滤���¼:" & chk_Ѫ͸����д�°滤���¼
            
    strTmp = strTmp & _
            "," & p�����ڲ����� & ":ǩ���Զ�λ��:" & chk_ǩ����λ & _
            "," & p�����ڲ����� & ":��ʾ��ǩλ��:" & chk_��ʾ��ǩλ�� & _
            "," & p�����ڲ����� & ":��ǩ��������Ϊǰ׺����:" & chk_ǩ������ǰ� & _
            "," & p�����ڲ����� & ":SyncPage:" & chk_�������ͬ����ҳ & _
            "," & p�����ڲ����� & ":ǩ��ʹ��ԭͼ:" & chk_ǩ��ʹ��ԭͼ & _
            "," & pסԺ�������� & ":ת�ƺ�Ҫ����д�Ĺ���������һҳ��ӡ:" & chk_ת�ƺ���д�Ĳ�������һҳ��ӡ & _
            "," & pסԺ�������� & ":�Զ���ʾ�������:" & chk_סԺ�����Զ���ʾ������� & _
            "," & pסԺ�������� & ":��������������д��������:" & chk_Ҫ������д�������� & _
            "," & p���Ӳ������ & ":���ղ��ܹ鵵:" & chk_�������������ܹ鵵 & _
            "," & p���Ӳ������ & ":��������¼��������:" & chk_��������¼�������� & _
            "," & p���Ӳ������� & ":����¼�����ԭ��:" & chk_����¼������������� & _
            "," & p���Ӳ������� & ":��������¼�����ԭ��:" & chk_��������¼�����ԭ�� & _
            "," & 0 & ":90:" & chk_�����Զ�д�벡����ҳ & _
            "," & 0 & ":91:" & chk_���������ȱ�Ŀ������

    strTmp = strTmp & _
            "," & p������Һ���� & ":Ƥ����֤���:" & chk_������ҺƤ����֤��� & _
            "," & p������Һ���� & ":δ�շ����:" & chk_������Һδ�շ��������
    
    strTmp = strTmp & _
        "," & p�°�סԺ��ʿվ & ":��Ƭ���������:" & chk_��Ƭ���������
    Call SetParToControl(strTmp, mrsPar, chk)
    '����ͼ
    strTmp = p�����¼���� & ":����ͼ��ʾ����ʱ��:" & chk_����ͼ��ʾ����ʱ�� & _
            "," & p�����¼���� & ":����ͼģʽ:" & chk_����ͼģʽ & _
            "," & p�����¼���� & ":��¶�ߵ���ʾλ��:" & chk_��¶�ߵ���ʾλ�� & _
            "," & p�����¼���� & ":����ͼ��ʾ������:" & chk_����ͼ����ʾ������
    '��¼��
    strTmp = strTmp & _
            "," & p�����¼���� & ":��¼��������ʾ��ʽ:" & chk_������ʾ��ʽ & _
            "," & p�����¼���� & ":��ҳ����ֻ��ʾ�ڵ�һҳ:" & chk_ֻ�ڵ�ǰҳ��ʾ��ҳ���� & _
            "," & p�����¼���� & ":��Ӧ��ݻ����ļ�:" & chk_��Ӧ��ݻ����ļ� & _
            "," & p�����¼���� & ":�����ļ�ҳ�����:" & chk_�ļ�ҳ����� & _
            "," & p�����¼���� & ":��¼��ǩ������ʾ��ʽ:" & chk_ǩ������ʾ��ʽ & _
            "," & p�����¼���� & ":��������ͬ��:" & chk_��������ͬ��
    '���µ�
    strTmp = strTmp & _
            "," & p�����¼���� & ":���±�־���λ��:" & chk_�Զ���־������� & _
            "," & p�����¼���� & ":���±�־����40�̶���С������ʾ:" & chk_�Զ���־����40�̶���С��ʾ & _
            "," & p�����¼���� & ":�ٴ�����ֹͣǰ�α�ע:" & chk_�ٴ�����ֹͣǰ�α�ע & _
            "," & p�����¼���� & ":���±�־��˳��������:" & chk_�Զ���־��˳�������� & _
            "," & p�����¼���� & ":�������ڸ�ʽ:" & chk_�������ڸ�ʽ & _
            "," & p�����¼���� & ":��Ʊ�ʶ���Զ�ת��Ϊ��Ժ:" & chk_��Ʊ�־���Զ�תΪ��Ժ & _
            "," & p�����¼���� & ":����������14���Ժ�����ʾ:" & chk_��������������Ժ & _
            "," & p�����¼���� & ":ȫ�������ʾ¼��ʱ��:" & chk_ȫ�������ʾСʱ & _
            "," & p�����¼���� & ":���ܲ�����ʾ��������:" & chk_������Ŀ��ʾ�������� & _
            "," & p�����¼���� & ":���µ�����ӡ������:" & chk_����������� & _
            "," & p�����¼���� & ":���µ���ʾ��ʽ:" & chk_��ͼ�̶ȵ�����ʾ & _
            "," & p�����¼���� & ":���µ���ʾ���:" & chk_��ʾ�����Ϣ & _
            "," & p�����¼���� & ":Ӥ�����µ�����������ʾ0:" & chk_Ӥ��סԺ������0��ʼ���� & _
            "," & p�����¼���� & ":��ӡҽԺ����:" & chk_��ӡҽԺ���� & _
            "," & p�����¼���� & ":���µ�����ӡ����˵��:" & chk_�������˵����Ϣ & _
            "," & p�����¼���� & ":������Ŀ������ʾ:" & chk_������Ŀ������ʾ
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '2.����ComboBox�����
    strTmp = "0:30:" & cbo_������ҩ�ӿ� & _
            ",0:224:" & cmd_����������Դ & _
            "," & pסԺҽ��վ & ":������ʾ��ʽ:" & cbo_סԺҽ��վ�б���ʾ & _
            "," & pסԺҽ��վ & ":������ҳ��׼:" & cbo_סԺ��ҳ��׼ & _
            "," & p�����ڲ����� & ":ǩ��ʱ��:" & cbo_����ǩ����ʾʱ��
    '��¼��
    strTmp = strTmp & _
            "," & p�����¼���� & ":��¼����ǩģʽ:" & cbo_��ǩģʽ & _
            "," & p�����¼���� & ":��ʿ��ǩ������ʾģʽ:" & cbo_ǩ������ʾģʽ
    '���µ�
    strTmp = strTmp & _
            "," & p�����¼���� & ":δ��˵����ʾλ��:" & cbo_δ��˵����ʾλ�� & _
            "," & p�����¼���� & ":���²�����ʾ��ʽ:" & cbo_���²�����ʾ��ʽ & _
            "," & p�����¼���� & ":�����������������ʽ:" & cbo_������������ʾλ�� & _
            "," & p�����¼���� & ":����������:" & cbo_�������������ʾλ�� & _
            "," & p�����¼���� & ":��������ȱʡ��ʽ:" & cbo_��������ȱʡ��ʽ & _
            "," & p�����¼���� & ":���±�־�ָ���:" & cbo_��־˵����ʱ�����ӷ���
    Call SetParToControl(strTmp, mrsPar, cbo)
            
    '3.����UpDown�����
    strTmp = "0:5:" & ud_��¼ҽ��ʶ���� & _
            ",0:220:" & ud_ҽ��ִ����Ч���� & _
            ",0:223:" & ud_�����¿�ҽ�����
    
    '��¼��
    strTmp = strTmp & "," & p�����¼���� & ":����¼�뻤����������:" & ud_��¼������¼������
    '���µ�
    strTmp = strTmp & _
            "," & p�����¼���� & ":�������߹̶��������:" & ud_�������߹̶�������� & _
            "," & p�����¼���� & ":���¿�ʼʱ��:" & ud_���¿�ʼʱ�� & _
            "," & p�����¼���� & ":���±������:" & ud_���±��̶��������
    Call SetParToControl(strTmp, mrsPar, ud)     'mrsPar�洢�Ŀؼ�����txtUD
    
    '4.����TextBox�����
    strTmp = p����ҽ���´� & ":��Ѫ����ע������:" & txt_������Ѫ����ע������ & _
            "," & pסԺҽ���´� & ":��Ѫ����ע������:" & txt_סԺ��Ѫ����ע������ & _
            "," & p�ٴ�·��Ӧ�� & ":��ҩ�䷽�����޸ĵ���ҩζ������:" & txt_��ҩ�䷽�����޸ĵ���ҩζ������ & _
            "," & p�ٴ�·��Ӧ�� & ":·��ҽ�����ɳ�ǰ����:" & txt_����ʱҽ����������ǰʱ������ & _
            "," & p����·��Ӧ�� & ":��ҩ�䷽�����޸ĵ���ҩζ������:" & txt_������ҩ�䷽�����޸ĵ���ҩζ������ & _
            "," & p����·��Ӧ�� & ":·��ҽ�����ɳ�ǰ����:" & txt_��������ʱҽ����������ǰʱ������ & _
            "," & p����ҽ��վ & ":����ˢ�¼��:" & txt_����ҽ��վ�����б��Զ�ˢ������ & _
            "," & pסԺҽ��վ & ":������鷴������:" & txt_סԺҽ������������ & _
            "," & pסԺ��ʿվ & ":������鷴������:" & txt_סԺ��ʿ���������� & _
            "," & pסԺҽ������ & ":��������ҩҽ��������ʾ����:" & txt_��������ҩҽ��������ʾ���� & _
            "," & pסԺҽ������ & ":��������ҩҽ��������ʾ����:" & txt_��������ҩҽ��������ʾ����


    strTmp = strTmp & _
            "," & p�����ڲ����� & ":ǩ��ͼƬ�߶�:" & txt_ǩ��ʹ��ͼƬ�߶� & _
            "," & pסԺ�������� & ":�������۵���ʼ����:" & txt_������N���Զ��۵� & _
            "," & p���Ӳ������ & ":������������:" & txt_���Ӳ������ȱʡ���� & _
            "," & p���Ӳ������ & ":δ����ˢ��Ƶ��:" & txt_���Ӳ�����������ˢ�¼�� & _
            "," & p���Ӳ������� & ":������������:" & txt_���Ӳ�������ȱʡ���� & _
            "," & p���Ӳ������� & ":���������:" & txt_���Ӳ��������������
    
    strTmp = strTmp & _
            "," & p������Һ���� & ":ҽ��ˢ�¼��:" & txt_������Һ�Զ�ˢ�²��� & _
            "," & p������Һ���� & ":Ƥ��������ǰʱ��:" & txt_������ҺƤ������ & _
            "," & p������Һ���� & ":Ĭ�ϵ���:" & txt_������Һ���� & _
            "," & p������Һ���� & ":Ĭ�ϵ�ϵ��:" & txt_������Һ��ϵ�� & _
            "," & p������Һ���� & ":��Һ������ǰʱ��:" & txt_������Һ��Һ����
    
    '���µ�
    strTmp = strTmp & _
        "," & p�����¼���� & ":�������ע����:" & txt_������ע���� & _
        "," & p�����¼���� & ":���¸��Ժϸ����:" & txt_���¸��Ժϸ����
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '5.����ListBox�����
    strTmp = "0:80:" & lst_סԺ���ͻ������ & _
            ",0:86:" & lst_���﷢�ͻ������ & _
            "," & p����ҽ���´� & ":Ҫ�������������:" & lst_���﷢�ͼ����� & _
            "," & p����ҽ���´� & ":����Ϊͬһ���ݵ�ҽ�����:" & lst_���﷢��һ�ŵ������ & _
            "," & pסԺҽ���´� & ":Ҫ��������Ժ���:" & lst_סԺ�����Ժ���
    Call SetParToControl(strTmp, mrsPar, lst)
        
    strTmp = pסԺҽ������ & ":��ҩ���ջ�:" & lst_���ջصķ�ҩ����
    Call SetParToControl(strTmp, mrsPar, lst, 2)
    
    strTmp = pסԺҽ������ & ":����ǼǸ�ҩ;������:" & lst_ҩƷ�����ҩ;��
    Call SetParToControl(strTmp, mrsPar, lst, 3)
    
    '6.����OptionButton�����
    arrObj = Array(p����ҽ���´�, "���͵�������", opt���͵�������, _
                    p����ҽ���´�, "���͵��ݺŹ���", opt���͵��ݹ���, _
                    p����ҽ���´�, "����ҩ��ȱʡ��ҩĿ��", opt����Ŀ������, _
                    pסԺҽ���´�, "ҽ������ӡģʽ", optסԺҽ������ӡ, _
                    pסԺҽ���´�, "����ҩ��ȱʡ��ҩĿ��", opt����Ŀ��סԺ, _
                    pסԺҽ���´�, "����Ƥ�Խ������ҽ����������", optδƤ������ҽ��, _
                    pסԺҽ������, "�����ջز�����������", opt���ڷ����ջ�, _
                    pסԺҽ������, "��Ѫ���뵥��ӡģʽ", opt��Ѫ���뵥��ӡ, _
                    pסԺҽ������, "ת�ƺͳ�Ժ��ӡ", optת��������Ժҽ��, _
                    pסԺҽ������, "סԺ��ҩ����", opt��ҩ����, _
                    p�ٴ�·��Ӧ��, "·������ӡ����", opt·����ӡ����, _
                    p�ٴ�·��Ӧ��, "·����ÿҳ��ӡ������", optÿҳ·����ӡ����, _
                    p����·��Ӧ��, "·������ӡ����", opt����·����ӡ����, _
                    p����·��Ӧ��, "·����ÿҳ��ӡ������", opt����·����ӡ����, _
                    pסԺҽ��վ, "�����ж����", opt�����ж�, _
                    pסԺҽ��վ, "������ϼ��", opt�������, _
                    pסԺҽ��վ, "ICD������", optICD����, _
                    pסԺҽ��վ, "������", opt����, _
                    pҽ������վ, "���˹��˷�ʽ", opt���˹���, _
                    p�����ڲ�����, "SignShow", optSign, _
                    p�°�סԺ��ʿվ, "��λ��Ƭ����ʽ", optNewCard, _
                    pסԺҽ������, "ҩƷ�÷�������ӡһ��", optPrintDruUse)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�ٴ�·��Ӧ��, "·������ӡ��ʽ", optPrintWay)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p����ҽ��վ, "���˽������", opt�������)
    Call SetParToControl("", mrsPar, arrObj, 1)
    '���µ�
    arrObj = Array(p�����¼����, "���������䷽ʽ", optPloy, _
                    p�����¼����, "���������(����/����)��ʽ¼��", OptInsert, _
                    p�����¼����, "��Ժ��־������ʾ", OptOut, _
                    p�����¼����, "���µ��ļ���ʼʱ��", optFileTime, _
                    p�����¼����, "�೦������ʾ��ʽ", OptEnemaStool)
    Call SetParToControl("", mrsPar, arrObj)
    
    '7.����ϵͳ����
    rsTmp.Filter = "ģ��=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case 56
            ud(ud_���ﴦ����������).value = IIF(Val(strValue) = 0, 5, Val(strValue))
            
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")   '����CheckBox�ؼ���������Ҫ�ٲ���һ����¼
            Call SetParRelation(txtUD, ud_���ﴦ����������, mrsPar)
                        
        Case 70
            ud(ud_�����Ǽ���Ч����).value = IIF(Val(strValue) = 0, 1, Val(strValue))
            
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "") '����CheckBox�ؼ���������Ҫ�ٲ���һ����¼
            Call SetParRelation(txtUD, ud_�����Ǽ���Ч����, mrsPar)
        Case 186
            chk(chk_��Ѫҽ��ִ�к���Ҫ�˶�).value = Mid(strValue, 1, 1)
            chk(chk_Ƥ��ҽ��ִ�к���Ҫ�˶�).value = Mid(strValue, 2, 1)
            
            Call SetParRelation(chk, chk_��Ѫҽ��ִ�к���Ҫ�˶�, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_Ƥ��ҽ��ִ�к���Ҫ�˶�, mrsPar)
        Case 188
            chk(chk_����ҩ��ʹ���Ա�ҩ).Enabled = chk(chk_����ҩ��ּ�����).value = 1
        Case 248
            chk(chk_��ҽ��С����п���ҩ�����).Enabled = chk(chk_����ҩ��ּ�����).value = 1
        Case 213
            cbo(cbo_��ҩ�䷽).ListIndex = IIF(Val(strValue) = 4, 1, 0)
            Call SetParRelation(cbo, cbo_��ҩ�䷽, mrsPar, rsTmp!������)
        Case 220    '����ȡ��n���ڵǼǵ�ҽ��ִ�м�¼
            chk(chk_ҽ��ִ����Ч����).value = IIF(Val(strValue) = 999, 0, 1)
            
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "") '����txtUD�ؼ���������Ҫ�ٲ���һ����¼
            Call SetParRelation(chk, chk_ҽ��ִ����Ч����, mrsPar)
            
        Case 228   '����������ҩ�ӿڰ汾
            strTmp = NVL(strValue, "3.0")
            If strTmp = "3.0" Then
                optPASSVer(0).value = True
            Else
                optPASSVer(1).value = True
            End If
            If Not mobjPass Is Nothing Then
                cmdSet.Visible = mobjPass.SetEnabled(cbo(cbo_������ҩ�ӿ�).ListIndex, strTmp)
            End If
            Call SetParRelation(optPASSVer, 0, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(optPASSVer, 1, mrsPar)
        Case 230
            Call SetVsfEditable(vsUnWriteDept, Val(strValue) = 1)
        Case 271
            Call SetVsfEditable(vsStopDept, Val(strValue) = 1)
        Case 273
            If mblnUseBlood Then
                Call zlControl.cbo.Locate(cbo(cbo_��Ѫ�ɼ�Ĭ����������), strValue)
                Call SetParRelation(cbo, cbo_��Ѫ�ɼ�Ĭ����������, mrsPar, rsTmp!������)
            End If
        Case 233
            Call Load����(vsUnWriteDept, strValue)
            Call SetParRelation(vsUnWriteDept, 0, mrsPar, rsTmp!������)
        Case 285
            Call Load����(vsStopDept, strValue)
            Call SetParRelation(vsStopDept, 0, mrsPar, rsTmp!������)
        Case 238
            Call Set���뵥���û���(strValue)
            Call SetParRelation(chk, chk_���뵥���û���������, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_���뵥���û����������, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_���뵥���û���������Ѫ, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_���뵥���û�����������, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_���뵥���û���סԺ���, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_���뵥���û���סԺ����, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_���뵥���û���סԺ��Ѫ, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_���뵥���û���סԺ����, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_���뵥���û��ڻ���, mrsPar)
        Case 260
            chk(chk_�������뵥�����ʹ�����뵥�´�ҽ������).value = Mid(strValue, 1, 1)
            chk(chk_�������뵥�����ʹ�����뵥�´�ҽ��סԺ).value = Mid(strValue, 2, 1)
            Call SetParRelation(chk, chk_�������뵥�����ʹ�����뵥�´�ҽ������, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_�������뵥�����ʹ�����뵥�´�ҽ��סԺ, mrsPar)
        End Select
        
        rsTmp.MoveNext
    Loop
    
    '8.����ģ���������
    rsTmp.Filter = "ģ��=" & p����ҽ���´�
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
            Case "ҩ�����Ҷ��շ���"
                'ҩ�����Ҷ��շ���
                For i = 0 To UBound(Split(strValue, ";"))
                    cbo(cbo_����ҩ�����Ҷ��շ���).AddItem "����" & i + 1
                    cbo(cbo_����ҩ�����Ҷ��շ���).ItemData(cbo(cbo_����ҩ�����Ҷ��շ���).NewIndex) = i + 1
                    mrs����ҩ������.AddNew
                    mrs����ҩ������!����IDs = Split(strValue, ";")(i)
                    mrs����ҩ������!���� = i + 1
                    mrs����ҩ������!ȱʡ��ҩ�� = zlDatabase.GetPara("����ȱʡ��ҩ��", glngSys, p����ҽ���´�, , , , , Val(Split(mrs����ҩ������!����IDs, ",")(0)))
                    mrs����ҩ������!ȱʡ��ҩ�� = zlDatabase.GetPara("����ȱʡ��ҩ��", glngSys, p����ҽ���´�, , , , , Val(Split(mrs����ҩ������!����IDs, ",")(0)))
                    mrs����ҩ������!ȱʡ��ҩ�� = zlDatabase.GetPara("����ȱʡ��ҩ��", glngSys, p����ҽ���´�, , , , , Val(Split(mrs����ҩ������!����IDs, ",")(0)))
                    mrs����ҩ������!������ҩ�� = zlDatabase.GetPara("���������ҩ��", glngSys, p����ҽ���´�, , , , , Val(Split(mrs����ҩ������!����IDs, ",")(0)))
                    mrs����ҩ������!���ó�ҩ�� = zlDatabase.GetPara("������ó�ҩ��", glngSys, p����ҽ���´�, , , , , Val(Split(mrs����ҩ������!����IDs, ",")(0)))
                    mrs����ҩ������!������ҩ�� = zlDatabase.GetPara("���������ҩ��", glngSys, p����ҽ���´�, , , , , Val(Split(mrs����ҩ������!����IDs, ",")(0)))
                    mrs����ҩ������!ȱʡ���ϲ��� = zlDatabase.GetPara("����ȱʡ���ϲ���", glngSys, p����ҽ���´�, , , , , Val(Split(mrs����ҩ������!����IDs, ",")(0)))
                    mrs����ҩ������!���÷��ϲ��� = zlDatabase.GetPara("������÷��ϲ���", glngSys, p����ҽ���´�, , , , , Val(Split(mrs����ҩ������!����IDs, ",")(0)))
                    mrs����ҩ������.Update
                Next
                If cbo(cbo_����ҩ�����Ҷ��շ���).ListCount > 0 Then
                    cbo(cbo_����ҩ�����Ҷ��շ���).ListIndex = 0
                End If
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsUseDept, cbo_����ҩ�����Ҷ��շ���, mrsPar)
            Case "����ȱʡ��ҩ��"
                '����ȱʡ��ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_����ҩ�����Ҷ��շ���, mrsPar, , , "����ȱʡ��ҩ��")
            Case "����ȱʡ��ҩ��"
                '����ȱʡ��ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_����ҩ�����Ҷ��շ���, mrsPar, , , "����ȱʡ��ҩ��")
            Case "����ȱʡ��ҩ��"
                '����ȱʡ��ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_����ҩ�����Ҷ��շ���, mrsPar, , , "����ȱʡ��ҩ��")
            Case "����ȱʡ���ϲ���"
                '����ȱʡ���ϲ���
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_����ҩ�����Ҷ��շ���, mrsPar, , , "����ȱʡ���ϲ���")
            Case "���������ҩ��"
                '���������ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_����ҩ�����Ҷ��շ���, mrsPar, , , "���������ҩ��")
            Case "������ó�ҩ��"
                '������ó�ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_����ҩ�����Ҷ��շ���, mrsPar, , , "������ó�ҩ��")
            Case "���������ҩ��"
                '���������ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_����ҩ�����Ҷ��շ���, mrsPar, , , "���������ҩ��")
            Case "������÷��ϲ���"
                '������÷��ϲ���
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_����ҩ�����Ҷ��շ���, mrsPar, , , "������÷��ϲ���")
        End Select
        
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "ģ��=" & pסԺҽ���´�
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
            Case "ҩ�����Ҷ��շ���"
                'ҩ�����Ҷ��շ���
                For i = 0 To UBound(Split(strValue, ";"))
                    cbo(cbo_סԺҩ�����Ҷ��շ���).AddItem "����" & i + 1
                    cbo(cbo_סԺҩ�����Ҷ��շ���).ItemData(cbo(cbo_סԺҩ�����Ҷ��շ���).NewIndex) = i + 1
                    mrsסԺҩ������.AddNew
                    mrsסԺҩ������!����IDs = Split(strValue, ";")(i)
                    mrsסԺҩ������!���� = i + 1
                    mrsסԺҩ������!ȱʡ��ҩ�� = zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , Val(Split(mrsסԺҩ������!����IDs, ",")(0)))
                    mrsסԺҩ������!ȱʡ��ҩ�� = zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , Val(Split(mrsסԺҩ������!����IDs, ",")(0)))
                    mrsסԺҩ������!ȱʡ��ҩ�� = zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , Val(Split(mrsסԺҩ������!����IDs, ",")(0)))
                    mrsסԺҩ������!������ҩ�� = zlDatabase.GetPara("סԺ������ҩ��", glngSys, pסԺҽ���´�, , , , , Val(Split(mrsסԺҩ������!����IDs, ",")(0)))
                    mrsסԺҩ������!���ó�ҩ�� = zlDatabase.GetPara("סԺ���ó�ҩ��", glngSys, pסԺҽ���´�, , , , , Val(Split(mrsסԺҩ������!����IDs, ",")(0)))
                    mrsסԺҩ������!������ҩ�� = zlDatabase.GetPara("סԺ������ҩ��", glngSys, pסԺҽ���´�, , , , , Val(Split(mrsסԺҩ������!����IDs, ",")(0)))
                    mrsסԺҩ������!ȱʡ���ϲ��� = zlDatabase.GetPara("סԺȱʡ���ϲ���", glngSys, pסԺҽ���´�, , , , , Val(Split(mrsסԺҩ������!����IDs, ",")(0)))
                    mrsסԺҩ������!���÷��ϲ��� = zlDatabase.GetPara("סԺ���÷��ϲ���", glngSys, pסԺҽ���´�, , , , , Val(Split(mrsסԺҩ������!����IDs, ",")(0)))
                    mrsסԺҩ������.Update
                Next
                If cbo(cbo_סԺҩ�����Ҷ��շ���).ListCount > 0 Then
                    cbo(cbo_סԺҩ�����Ҷ��շ���).ListIndex = 0
                End If
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsUseDept, cbo_סԺҩ�����Ҷ��շ���, mrsPar)
            Case "סԺȱʡ��ҩ��"
                'סԺȱʡ��ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_סԺҩ�����Ҷ��շ���, mrsPar, , , "סԺȱʡ��ҩ��")
            Case "סԺȱʡ��ҩ��"
                'סԺȱʡ��ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_סԺҩ�����Ҷ��շ���, mrsPar, , , "סԺȱʡ��ҩ��")
            Case "סԺȱʡ��ҩ��"
                'סԺȱʡ��ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_סԺҩ�����Ҷ��շ���, mrsPar, , , "סԺȱʡ��ҩ��")
            Case "סԺȱʡ���ϲ���"
                'סԺȱʡ���ϲ���
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_סԺҩ�����Ҷ��շ���, mrsPar, , , "סԺȱʡ���ϲ���")
            Case "סԺ������ҩ��"
                'סԺ������ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_סԺҩ�����Ҷ��շ���, mrsPar, , , "סԺ������ҩ��")
            Case "סԺ���ó�ҩ��"
                'סԺ���ó�ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_סԺҩ�����Ҷ��շ���, mrsPar, , , "סԺ���ó�ҩ��")
            Case "סԺ������ҩ��"
                'סԺ������ҩ��
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_סԺҩ�����Ҷ��շ���, mrsPar, , , "סԺ������ҩ��")
            Case "סԺ���÷��ϲ���"
                'סԺ���÷��ϲ���
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_סԺҩ�����Ҷ��շ���, mrsPar, , , "סԺ���÷��ϲ���")
        End Select
        
        rsTmp.MoveNext
    Loop
    
    
    rsTmp.Filter = "ģ��=" & pסԺҽ������
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "����ִ���Զ����"
            chk(chk_סԺ�����Զ�ִ�г���).value = Mid(strValue, 1, 1)
            chk(chk_סԺ�����Զ�ִ������).value = Mid(strValue, 2, 1)
            
            Call SetParRelation(chk, chk_סԺ�����Զ�ִ�г���, mrsPar, rsTmp!������, pסԺҽ������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_סԺ�����Զ�ִ������, mrsPar)
        
        Case "�����ڷ�ҩ���ͽ���ʱ��"
            If InStr(strValue, "|") = 0 Then
                chk(chk_�����ڷ�ҩ���ͽ���ʱ��).value = 0
            Else
                chk(chk_�����ڷ�ҩ���ͽ���ʱ��).value = Val(Split(strValue, "|")(0))
                If chk(chk_�����ڷ�ҩ���ͽ���ʱ��).value = 1 Then
                    dtp�ڷ�����ʱ��.value = Format(Split(strValue, "|")(1), "HH:MM:SS")
                    dtp�ڷ�����ʱ��.Enabled = True
                End If
            End If
            
            Call SetParRelation(chk, chk_�����ڷ�ҩ���ͽ���ʱ��, mrsPar, rsTmp!������, pסԺҽ������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(dtp�ڷ�����ʱ��, 0, mrsPar)
        
        Case "סԺ����ִ���Զ���ɷ���"
                'סԺ����ִ���Զ���ɷ���
                For i = 0 To UBound(Split(strValue, ";"))
                    cbo(cbo_סԺ����ִ���Զ���ɷ���).AddItem "����" & i + 1
                    cbo(cbo_סԺ����ִ���Զ���ɷ���).ItemData(cbo(cbo_סԺ����ִ���Զ���ɷ���).NewIndex) = i + 1
                    mrsסԺִ�ж���.AddNew
                    mrsסԺִ�ж���!����IDs = Split(strValue, ";")(i)
                    mrsסԺִ�ж���!���� = i + 1
                    mrsסԺִ�ж���!ҽ����� = zlDatabase.GetPara("����ִ���Զ����ҽ�����", glngSys, pסԺҽ������, , , , , Val(Split(mrsסԺִ�ж���!����IDs, ",")(0)))
                    mrsסԺִ�ж���.Update
                Next
                If cbo(cbo_סԺ����ִ���Զ���ɷ���).ListCount > 0 Then
                    cbo(cbo_סԺ����ִ���Զ���ɷ���).ListIndex = 0
                End If
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsUseDept, cbo_סԺ����ִ���Զ���ɷ���, mrsPar)
            Case "����ִ���Զ����ҽ�����"
                '����ִ���Զ����ҽ�����
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(lst, lst_����ִ���Զ����ҽ�����, mrsPar, , , "����ִ���Զ����ҽ�����")
                
        End Select
        
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "ģ��=" & p�ٴ�·��Ӧ��
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "·��ִ�л������ó���"
            chk(chk_·��ִ�л���ҽ������).value = Mid(strValue, 1, 1)
            chk(chk_·��ִ�л��ڻ�ʿ����).value = Mid(strValue, 2, 1)
            
            Call SetParRelation(chk, chk_·��ִ�л���ҽ������, mrsPar, rsTmp!������, p�ٴ�·��Ӧ��)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_·��ִ�л��ڻ�ʿ����, mrsPar)
        End Select
        
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "ģ��=" & p����ҽ��վ
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "���˽������"
            txt(txt_������ǰ�������) = Mid(strValue, 3)
        End Select
        
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "ģ��=" & pסԺ��������
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
            Case "����������Ԥ��"
                If strValue = -1 Then
                    optEprRead(0).value = True
                ElseIf strValue = 0 Then
                    optEprRead(1).value = True
                Else
                    optEprRead(2).value = True
                    txt(txt_����������Ԥ��N��).Text = strValue
                End If
                Call SetParRelation(optEprRead, 0, mrsPar, rsTmp!������, pסԺ��������)
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(optEprRead, 1, mrsPar)
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(optEprRead, 2, mrsPar)
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(txt, txt_����������Ԥ��N��, mrsPar)
        End Select
        rsTmp.MoveNext
    Loop
    
    'ʹ��ͼƬǩ������
    Call setDepartSign
    
    rsTmp.Filter = "ģ��=" & p���Ӳ������
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
            Case "��������˳��"
                If strValue = "" Then strValue = "5;1;6;2;3;4;8;7;9"
                arrTmp = Split(strValue, ";")
                With vsfEpr(0)
                    '1-סԺҽ��;2-סԺ����;3-������;4-�����¼;5-��ҳ��¼;6-ҽ������;7-����֤��;8-֪���ļ�
                    For i = 0 To UBound(arrTmp)
                        Select Case arrTmp(i)
                        Case "1"
                            .TextMatrix(i + 1, 0) = "סԺҽ��"
                            .RowData(i + 1) = 1
                        Case "2"
                            .TextMatrix(i + 1, 0) = "סԺ����"
                            .RowData(i + 1) = 2
                        Case "3"
                            .TextMatrix(i + 1, 0) = "������"
                            .RowData(i + 1) = 3
                        Case "4"
                            .TextMatrix(i + 1, 0) = "�����¼"
                            .RowData(i + 1) = 4
                        Case "5"
                            .TextMatrix(i + 1, 0) = "��ҳ��¼"
                            .RowData(i + 1) = 5
                        Case "6"
                            .TextMatrix(i + 1, 0) = "ҽ������"
                            .RowData(i + 1) = 6
                        Case "7"
                            .TextMatrix(i + 1, 0) = "����֤��"
                            .RowData(i + 1) = 7
                        Case "8"
                            .TextMatrix(i + 1, 0) = "֪���ļ�"
                            .RowData(i + 1) = 8
                        Case "9"
                            .TextMatrix(i + 1, 0) = "�ٴ�·��"
                            .RowData(i + 1) = 9
                        End Select
                    Next
                End With
                Call SetParRelation(vsfEpr, 0, mrsPar, rsTmp!������, p���Ӳ������)
            Case "�����ҷ�Χ"
                '���ֵ
                gstrSQL = "Select ID,���,���� From ��Ա��"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                
                gstrSQL = "Select a.ID,a.����,a.���� From ���ű� a,��������˵�� b Where a.ID=b.����id And b.��������='�ٴ�' and ( TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null)"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                
                With vsfEpr(1)
                    .Rows = 2
                    arrTmp = Split(strValue, ";")
                    For i = 0 To UBound(arrTmp)
                        varTmp = Split(arrTmp(i), ",") '����ID
                        rs.Filter = ""
                        rs.Filter = "ID=" & Val(varTmp(0))
                        If rs.RecordCount > 0 Then '
                            If Val(.TextMatrix(.Rows - 1, 0)) > 0 Then .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 1) = rs!����
                            .TextMatrix(.Rows - 1, 0) = rs!ID
                            
                            For n = 1 To UBound(varTmp)
                                rsData.Filter = ""
                                rsData.Filter = "ID=" & Val(varTmp(n))
                                If rsData.RecordCount > 0 Then
                                    If .TextMatrix(.Rows - 1, 3) = "" Then
                                        .TextMatrix(.Rows - 1, 3) = rsData!����
                                        .TextMatrix(.Rows - 1, 2) = rsData!ID
                                    Else
                                        .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 1, 3) & vbCrLf & rsData!����
                                        .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & rsData!ID
                                    End If
                                End If
                            Next
                        End If
                    Next
                    .Rows = .Rows + 1
                    .AutoSize 3, 3
                End With
                
                Call SetParRelation(vsfEpr, 1, mrsPar, rsTmp!������, p���Ӳ������)
        End Select
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "ģ��=" & p���Ӳ�������
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
            Case "���ֿ��ҷ�Χ"
                '���ֵ
                gstrSQL = "Select ID,���,���� From ��Ա��"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                
                gstrSQL = "Select a.ID,a.����,a.���� From ���ű� a,��������˵�� b Where a.ID=b.����id And b.��������='�ٴ�' and ( TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null)"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                
                With vsfEpr(2)
                    .Rows = 2
                    arrTmp = Split(strValue, ";")
                    For i = 0 To UBound(arrTmp)
                        varTmp = Split(arrTmp(i), ",") '����ID
                        rs.Filter = ""
                        rs.Filter = "ID=" & Val(varTmp(0))
                        If rs.RecordCount > 0 Then '
                            If Val(.TextMatrix(.Rows - 1, 0)) > 0 Then .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 1) = rs!����
                            .TextMatrix(.Rows - 1, 0) = rs!ID
                            
                            For n = 1 To UBound(varTmp)
                                rsData.Filter = ""
                                rsData.Filter = "ID=" & Val(varTmp(n))
                                If rsData.RecordCount > 0 Then
                                    If .TextMatrix(.Rows - 1, 3) = "" Then
                                        .TextMatrix(.Rows - 1, 3) = rsData!����
                                        .TextMatrix(.Rows - 1, 2) = rsData!ID
                                    Else
                                        .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 1, 3) & vbCrLf & rsData!����
                                        .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & rsData!ID
                                    End If
                                End If
                            Next
                        End If
                    Next
                    .Rows = .Rows + 1
                    .AutoSize 3, 3
                End With
            
                Call SetParRelation(vsfEpr, 2, mrsPar, rsTmp!������, p���Ӳ�������)
        End Select
        rsTmp.MoveNext
    Loop
    
    '������Һ����
    Call SetWaittingMixDept
    rsTmp.Filter = "ģ��=" & p������Һ����
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "����Һ�����б�"
            Call SetParRelation(vsfWaittingMixDept, 0, mrsPar, rsTmp!������, p������Һ����)
        End Select
        rsTmp.MoveNext
    Loop
    
    '�����¼��Ŀ����
    chk(chk_��������ͬ��).Enabled = chk(chk_��Ӧ��ݻ����ļ�).value = 1           '96044,����,����ͬ��Ϊ��Ӧ��ݻ����ļ����Ӳ���
    rsTmp.Filter = "ģ��=" & p�����¼����
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "���µ����" '���µ����
            arrTmp = Split(strValue, ";")
            If UBound(arrTmp) >= 0 Then cbo(cbo_��Ժ�Զ���־).ListIndex = Val(arrTmp(0))
            If UBound(arrTmp) >= 1 Then cbo(cbo_����Զ���־).ListIndex = Val(arrTmp(1))
            If UBound(arrTmp) >= 2 Then cbo(cbo_ת���Զ���־).ListIndex = Val(arrTmp(2))
            If UBound(arrTmp) >= 3 Then cbo(cbo_�����Զ���־).ListIndex = Val(arrTmp(3))
            If UBound(arrTmp) >= 4 Then cbo(cbo_�����Զ���־).ListIndex = Val(arrTmp(4))
            If UBound(arrTmp) >= 5 Then cbo(cbo_��Ժ�Զ���־).ListIndex = Val(arrTmp(5))
            If UBound(arrTmp) >= 6 Then cbo(cbo_�����Զ���־).ListIndex = Val(arrTmp(6))
            If UBound(arrTmp) >= 7 Then cbo(cbo_�����Զ���־).ListIndex = Val(arrTmp(7))
            If UBound(arrTmp) >= 8 Then cbo(cbo_�����Զ���־).ListIndex = Val(arrTmp(8))
            If UBound(arrTmp) >= 9 Then cbo(cbo_ת�����Զ���־).ListIndex = Val(arrTmp(9))
            Call SetParRelation(cbo, cbo_��Ժ�Զ���־, mrsPar, rsTmp!������, p�����¼����)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_����Զ���־, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_ת���Զ���־, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_�����Զ���־, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_�����Զ���־, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_��Ժ�Զ���־, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_�����Զ���־, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_�����Զ���־, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_�����Զ���־, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_ת�����Զ���־, mrsPar)
        Case "����ʱ��ҹ���־" '����ʱ��ҹ���־
            arrTmp = Split(strValue, ";")
            If UBound(arrTmp) >= 1 Then
                ud(ud_����ҹ�࿪ʼʱ��).value = Abs(Val(arrTmp(0)))
                ud(ud_����ҹ�����ʱ��).value = Abs(Val(arrTmp(1)))
            Else
                ud(ud_����ҹ�࿪ʼʱ��).value = Abs(Val(strValue))
            End If
            Call SetParRelation(txtUD, txtud_����ҹ�࿪ʼʱ��, mrsPar, rsTmp!������, p�����¼����)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(txtUD, txtud_����ҹ�����ʱ��, mrsPar)
        Case "�����������߱�־" '���߱�־(˳��)
            arrTmp = Split(strValue, ";")
            For i = 0 To 1
                intIndex = IIF(i = 0, cbo_��������˳��, cbo_��¶�½�˳��)
                If UBound(arrTmp) >= i Then
                    lngValue = Val(arrTmp(i))
                    If lngValue < 0 Or lngValue > cbo(intIndex).ListCount - 1 Then lngValue = 0
                    cbo(intIndex).ListIndex = lngValue
                Else
                    cbo(intIndex).ListIndex = 0
                End If
            Next
            
            Call SetParRelation(cbo, cbo_��������˳��, mrsPar, rsTmp!������, p�����¼����)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_��¶�½�˳��, mrsPar)
        Case "����������ʩ��־" '��ʩ��־
            arrTmp = Split(strValue, ";")
            For i = 0 To 1
                intIndex = IIF(i = 0, cbo_������־����, cbo_������־λ��)
                If UBound(arrTmp) >= i Then
                    lngValue = Val(arrTmp(i))
                    If lngValue < 0 Or lngValue > cbo(intIndex).ListCount - 1 Then lngValue = 0
                    cbo(intIndex).ListIndex = lngValue
                Else
                    cbo(intIndex).ListIndex = 0
                End If
            Next
            
            Call SetParRelation(cbo, cbo_������־����, mrsPar, rsTmp!������, p�����¼����)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_������־λ��, mrsPar)
        Case "���̾����쳣�߱�־" '�����ߺ��쳣�߱�־
            arrTmp = Split(strValue, ";")
            For i = 0 To 1
                intIndex = IIF(i = 0, cbo_��������ʾ, cbo_�쳣����ʾ)
                If UBound(arrTmp) >= i Then
                    lngValue = Val(arrTmp(i))
                    If lngValue < 0 Or lngValue > cbo(intIndex).ListCount - 1 Then lngValue = 0
                    cbo(intIndex).ListIndex = lngValue
                Else
                    cbo(intIndex).ListIndex = 0
                End If
            Next
            
            Call SetParRelation(cbo, cbo_��������ʾ, mrsPar, rsTmp!������, p�����¼����)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_�쳣����ʾ, mrsPar)
        Case "�����������߱�־(��)" '���߱�־(�쳣��)
            arrTmp = Split(strValue, ";")
            For i = 0 To 1
                intIndex = IIF(i = 0, cbo_���������쳣��, cbo_��¶�½��쳣��)
                If UBound(arrTmp) >= i Then
                    lngValue = Val(arrTmp(i))
                    If lngValue < 0 Or lngValue > cbo(intIndex).ListCount - 1 Then lngValue = 0
                    cbo(intIndex).ListIndex = lngValue
                Else
                    cbo(intIndex).ListIndex = 0
                End If
            Next
            
            Call SetParRelation(cbo, cbo_���������쳣��, mrsPar, rsTmp!������, p�����¼����)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_��¶�½��쳣��, mrsPar)
        Case "�������ߵ���0������" '0�����״ε�����
            lngValue = Val(strValue)
            If lngValue < 0 Or lngValue > cbo(cbo_������״ε�����).ListCount - 1 Then lngValue = 0
            cbo(cbo_������״ε�����).ListIndex = lngValue
            Call SetParRelation(cbo, cbo_������״ε�����, mrsPar, rsTmp!������, p�����¼����)
        Case "С���ʶ��ɫ" 'С���ʶ��ɫ
            lngValue = Val(strValue)
            picLineColor(0).BackColor = lngValue
            Call SetParRelation(picLineColor, 0, mrsPar, rsTmp!������, p�����¼����)
        Case "���µ������ʾ��ɫ" '���±�־��ɫ
            lngValue = Val(strValue)
             picLineColor(1).BackColor = lngValue
            Call SetParRelation(picLineColor, 1, mrsPar, rsTmp!������, p�����¼����)
        Case "δ��˵����ʾ��ɫ" 'δ��˵����ɫ
            lngValue = Val(strValue)
             picLineColor(2).BackColor = lngValue
            Call SetParRelation(picLineColor, 2, mrsPar, rsTmp!������, p�����¼����)
        Case "����������ʾ��ɫ" '����������ɫ
            lngValue = Val(strValue)
            picLineColor(3).BackColor = lngValue
            Call SetParRelation(picLineColor, 3, mrsPar, rsTmp!������, p�����¼����)
        Case "���¸��Ժϸ���ɫ" '���¸�����ɫ
            lngValue = Val(strValue)
            picLineColor(4).BackColor = lngValue
            Call SetParRelation(picLineColor, 4, mrsPar, rsTmp!������, p�����¼����)
        Case "С��ȱʡ��ʽ"
            arrTmp = Split(strValue, ";")
            If UBound(arrTmp) >= 0 Then cbo(cbo_С��ȱʡ��ʶ).ListIndex = Val(arrTmp(0))
            If UBound(arrTmp) >= 1 Then chk(chk_С��ȱʡ��ʶ).value = Val(arrTmp(1))
            Call SetParRelation(cbo, cbo_С��ȱʡ��ʶ, mrsPar, rsTmp!������, p�����¼����)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_С��ȱʡ��ʶ, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
        End Select
        rsTmp.MoveNext
    Loop
    '�°�סԺ��ʿվ
    txt(txt_���廤��IP��ַ) = ""
    txt(txt_���廤��IP�˿�) = ""
    rsTmp.Filter = "ģ��=" & p�°�סԺ��ʿվ
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
            Case "���廤��IP��ַ"
                If InStr(1, strValue, ":") <> 0 Then
                    txt(txt_���廤��IP��ַ) = Mid(strValue, 1, InStr(1, strValue, ":") - 1)
                    txt(txt_���廤��IP�˿�) = Mid(strValue, InStr(1, strValue, ":") + 1)
                Else
                    txt(txt_���廤��IP��ַ) = strValue
                End If
                Call SetParRelation(txt, txt_���廤��IP��ַ, mrsPar, rsTmp!������, p�°�סԺ��ʿվ)
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(txt, txt_���廤��IP�˿�, mrsPar)
        End Select
        rsTmp.MoveNext
    Loop
End Sub

Private Sub InitEnv()
'���ܣ���ʼ������ؼ������ػ�������
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strTmp As String
    Dim blnTmp As Boolean
    
    On Error GoTo ErrHandle

    vsUnWriteDept.ComboList = "..."
    vsUnWriteDept.RowHeightMin = 280
    vsStopDept.ComboList = "..."
    vsStopDept.RowHeightMin = 280
    vsUseDept(cbo_����ҩ�����Ҷ��շ���).ColWidth(0) = 3000
    vsUseDept(cbo_סԺҩ�����Ҷ��շ���).ColWidth(0) = 3000
    vsUseDept(cbo_סԺ����ִ���Զ���ɷ���).ColWidth(0) = 3000
    vsUseDept(cbo_����ҩ�����Ҷ��շ���).ColAlignment(0) = flexAlignLeftCenter
    vsUseDept(cbo_סԺҩ�����Ҷ��շ���).ColAlignment(0) = flexAlignLeftCenter
    vsUseDept(cbo_סԺ����ִ���Զ���ɷ���).ColAlignment(0) = flexAlignLeftCenter
    
    cbo(cbo_��ҩ�䷽).AddItem "0-��ζ��ҩ"
    cbo(cbo_��ҩ�䷽).AddItem "1-��ζ��ҩ"
    cbo(cbo_��ҩ�䷽).ListIndex = 0
    
    cbo(cbo_������ҩ�ӿ�).AddItem "0-δʹ��"
    cbo(cbo_������ҩ�ӿ�).AddItem "1-�Ĵ�����"
    cbo(cbo_������ҩ�ӿ�).AddItem "2-�Ϻ���ͨ"
    cbo(cbo_������ҩ�ӿ�).AddItem "3-����̫Ԫͨ"
    cbo(cbo_������ҩ�ӿ�).AddItem "4-���ݱ���"
    cbo(cbo_������ҩ�ӿ�).AddItem "5-��������"
    cbo(cbo_������ҩ�ӿ�).AddItem "6-������Ϣ"
    
    cbo(cbo_������ҩ�ӿ�).ListIndex = 0
    
    cbo(cmd_����������Դ).AddItem "0-��ѡ��������Դ"
    cbo(cmd_����������Դ).AddItem "1-��ҩƷĿ¼����"
    cbo(cmd_����������Դ).AddItem "2-������Դ����"
    cbo(cmd_����������Դ).ListIndex = 0
    
    cbo(cbo_סԺҽ��վ�б���ʾ).AddItem "����"
    cbo(cbo_סԺҽ��վ�б���ʾ).AddItem "����"
    cbo(cbo_סԺҽ��վ�б���ʾ).ListIndex = 0
    
    cbo(cbo_סԺ��ҳ��׼).AddItem "0-��������׼"
    cbo(cbo_סԺ��ҳ��׼).AddItem "1-�Ĵ�ʡ��׼"
    cbo(cbo_סԺ��ҳ��׼).AddItem "2-����ʡ��׼"
    cbo(cbo_סԺ��ҳ��׼).AddItem "3-����ʡ��׼"
    cbo(cbo_סԺ��ҳ��׼).ListIndex = 0
    
    '��ȡҽ�����ݶ���
    gstrSQL = "Select �������,ҽ������ From ҽ�����ݶ��� Order by �������"
    Call zlDatabase.OpenRecordset(mrsAdvice, gstrSQL, Me.Caption)
    
    
    '��ȡҽ������Ϊ�������
    gstrSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('5','6','7','8','9')" & _
        " Union All Select '5','ҩƷ' From Dual Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
  
    Do While Not rsTmp.EOF
        lst(lst_���﷢�ͻ������).AddItem rsTmp!���� & "-" & rsTmp!����
        lst(lst_���﷢�ͻ������).ItemData(lst(lst_���﷢�ͻ������).NewIndex) = Asc(rsTmp!����)
        
        lst(lst_סԺ���ͻ������).AddItem rsTmp!���� & "-" & rsTmp!����
        lst(lst_סԺ���ͻ������).ItemData(lst(lst_סԺ���ͻ������).NewIndex) = Asc(rsTmp!����)
        
        rsTmp.MoveNext
    Loop

    If rsTmp.RecordCount > 0 Then rsTmp.Filter = "����<>'4'"
    Do While Not rsTmp.EOF
        lst(lst_���﷢�ͼ�����).AddItem rsTmp!���� & "-" & rsTmp!����
        lst(lst_���﷢�ͼ�����).ItemData(lst(lst_���﷢�ͼ�����).NewIndex) = Asc(rsTmp!����)
        
        
        lst(lst_סԺ�����Ժ���).AddItem rsTmp!���� & "-" & rsTmp!����
        lst(lst_סԺ�����Ժ���).ItemData(lst(lst_סԺ�����Ժ���).NewIndex) = Asc(rsTmp!����)
        rsTmp.MoveNext
    Loop
    lst(lst_���﷢�ͼ�����).ListIndex = 0
    lst(lst_סԺ�����Ժ���).ListIndex = 0
    
    With lst(lst_���﷢��һ�ŵ������)
        If rsTmp.RecordCount > 0 Then rsTmp.Filter = "����<>'4' And ����<>'5'"
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & "-" & rsTmp!����
            .ItemData(.NewIndex) = Asc(rsTmp!����)
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    With lst(lst_����ִ���Զ����ҽ�����)
        .Clear
        .AddItem "��Һ"
        .ItemData(.NewIndex) = 0
        .AddItem "ע��"
        .ItemData(.NewIndex) = 1
        .AddItem "�ڷ�"
        .ItemData(.NewIndex) = 2
        .AddItem "�ɼ�"
        .ItemData(.NewIndex) = 3
        .AddItem "��������"
        .ItemData(.NewIndex) = 4
        .AddItem "��ͨ����"
        .ItemData(.NewIndex) = 5
        .AddItem "��������"
        .ItemData(.NewIndex) = 6
        .AddItem "������ҩ;��"
        .ItemData(.NewIndex) = 7
        .AddItem "����ҽ��"
        .ItemData(.NewIndex) = 8
        .ListIndex = 0
    End With
     
    '���ջصķ�ҩ����
    gstrSQL = "Select ����, ���� From ��ҩ���� Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    With lst(lst_���ջصķ�ҩ����)
                .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            rsTmp.MoveNext
        Loop
        If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0
    End With
    
    Call InitRsҩ������(mrs����ҩ������)
    Call InitRsҩ������(mrsסԺҩ������)
    Set mrsסԺִ�ж��� = New ADODB.Recordset
    mrsסԺִ�ж���.Fields.Append "����", adVarChar, 1000
    mrsסԺִ�ж���.Fields.Append "����IDs", adVarChar, 40000
    mrsסԺִ�ж���.Fields.Append "ҽ�����", adVarChar, 40000
    mrsסԺִ�ж���.CursorLocation = adUseClient
    mrsסԺִ�ж���.LockType = adLockOptimistic
    mrsסԺִ�ж���.CursorType = adOpenStatic
    mrsסԺִ�ж���.Open
    
    '��ҩҩ��
    InitVsDrugStore cbo_����ҩ�����Ҷ��շ���
    InitVsDrugStore cbo_סԺҩ�����Ҷ��շ���
    
    'ҩƷ���̵ǼǸ�ҩ;��
    gstrSQL = "Select ID,����,���� From ������ĿĿ¼" & _
        " Where ���='E' And ��������='2' And ������� IN(2,3) And (վ��='" & gstrNodeNo & "' Or վ�� is Null) Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    With lst(lst_ҩƷ�����ҩ;��)
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            .ItemData(.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0
    End With
    
    With vsfMecItem
        .Rows = 1
        .Cols = 3
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����"
        .ColWidth(0) = 1000
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .Cell(flexcpAlignment, 0, 0, 0, 2) = 4
    End With
    gstrSQL = "select ����,����,���� from ������Ŀ order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    cmdModify.Enabled = rsTmp.RecordCount > 0: cmdDelete.Enabled = rsTmp.RecordCount > 0
    While Not rsTmp.EOF
        With vsfMecItem
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsTmp!���� & ""
            .TextMatrix(.Rows - 1, 1) = rsTmp!���� & ""
            .TextMatrix(.Rows - 1, 2) = rsTmp!���� & ""
        End With
        rsTmp.MoveNext
    Wend
    
    With cbo(2)
        .Clear
        .AddItem "����ʾ"
        .AddItem Format(Now(), "yyyy-MM-dd hh:mm")
        .AddItem Format(Now(), "yyyy��MM��dd�� hh:mm")
    End With
    
    With vsfEpr(0)
        .Clear
        .Rows = 10
        .RowHidden(0) = True
    End With
    
    With vsfEpr(1)
        .Clear
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "����ID"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����ID"
        .TextMatrix(0, 3) = "����"
        .ColWidth(0) = 0
        .ColWidth(1) = 1600
        .ColWidth(2) = 0
        .ColWidth(3) = 2600
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Editable = flexEDKbd
    End With
    
    With vsfEpr(2)
        .Clear
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "����ID"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����ID"
        .TextMatrix(0, 3) = "����"
        .ColWidth(0) = 0
        .ColWidth(1) = 1600
        .ColWidth(2) = 0
        .ColWidth(3) = 2600
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Editable = flexEDKbd
    End With
    
    '�ж��Ƿ��ܹ������°�Ѫ��ϵͳ
    If mblnUseBlood Then
        '��ȡ���Ƽ�������
        gstrSQL = "Select ����,����,ȱʡ��־ From ���Ƽ������� order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���Ƽ�������")
        cbo(cbo_��Ѫ�ɼ�Ĭ����������).Clear
        i = 0
        Do While Not rsTmp.EOF
            cbo(cbo_��Ѫ�ɼ�Ĭ����������).AddItem rsTmp!���� & "-" & rsTmp!����
            If Val("" & rsTmp!ȱʡ��־) = 1 Then
                cbo(cbo_��Ѫ�ɼ�Ĭ����������).ListIndex = cbo(cbo_��Ѫ�ɼ�Ĭ����������).NewIndex
            End If
            If "" & rsTmp!���� = "Ѫ����" Then i = cbo(cbo_��Ѫ�ɼ�Ĭ����������).NewIndex
            rsTmp.MoveNext
        Loop
        If cbo(cbo_��Ѫ�ɼ�Ĭ����������).ListIndex = -1 And cbo(cbo_��Ѫ�ɼ�Ĭ����������).ListCount > 0 Then
            cbo(cbo_��Ѫ�ɼ�Ĭ����������).ListIndex = i
        End If
        cbo(cbo_��Ѫ�ɼ�Ĭ����������).Tag = cbo(cbo_��Ѫ�ɼ�Ĭ����������).ListIndex
        
        FrmBloodManager.Visible = True
    Else
        FrmBloodManager.Visible = False
        chk(chk_����Ѫ�����ϵͳ).Left = chk(chk_��ƻ��������дҪ��).Left
    End If
    
    Call InitNurseItem
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitVsDrugStore(ByVal intIndex As Long)
    Dim strSQL As String, rsTmp As Recordset
    Dim j As Long, i As Long
    Dim lngRow As Long, arrTmp As Variant
    
    'ҩ���뷢�ϲ���
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.�������� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " AND B.����ID=A.ID And B.������� IN(" & IIF(intIndex = cbo_����ҩ�����Ҷ��շ���, "1,3", "2,3") & ") and B.�������� in('��ҩ��','��ҩ��','��ҩ��','���ϲ���')" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by ��������,����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    With vsfDrugStore(intIndex)
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        .MergeCol(.ColIndex("���")) = True
        .MergeCells = flexMergeFixedOnly
        
        If intIndex = cbo_����ҩ�����Ҷ��շ��� Then
            '���� ��ҩ����  �У�1252 ģ����3������û���õ� '��ҩ������','��ҩ������','��ҩ������'  ����������
            .ColHidden(.ColIndex("��ҩ����")) = True
        End If
        
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = .FixedRows
            arrTmp = Split("��ҩ��,��ҩ��,��ҩ��,���ϲ���", ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "��������='" & arrTmp(i) & "'"
                Do While Not rsTmp.EOF
                    .TextMatrix(lngRow, .ColIndex("���")) = arrTmp(i)
                    .TextMatrix(lngRow, .ColIndex("ҩ��")) = rsTmp!����
                    .RowData(lngRow) = Val(rsTmp!ID)
                    
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Loop
                If lngRow < .Rows - 1 Then  '���ָ���
                    .Select lngRow, .FixedCols, lngRow, .Cols - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdOK_Click()
    Dim lngTmp As Long
    
    If ValidateData() = False Then Exit Sub
    
    Call Saveҽ������
    Call Saveҩ�����Ҷ���
    Call SaveDepartSign
    Call SaveסԺִ�ж���
    lngTmp = -1
    mrsPar.Filter = "������='����ҩ�����հ�������������' and �޸�״̬=1"
    If Not mrsPar.EOF Then
        lngTmp = mrsPar!������ֵ
    End If
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    
    Call Setҩ�����Ҷ��ղ�������(lngTmp)
    
    mblnOk = True
    Unload Me
End Sub

Private Function ValidateData() As Boolean
'���ܣ���֤���ݵ���Ч��
    mrsPar.Filter = "ģ��=" & p�°�סԺ��ʿվ & " and ������='���廤��IP��ַ' and �޸�״̬=1"
    If cmdLink.Tag <> "OK" And Not mrsPar.EOF Then
        If CheckNurseIntegrateIP(False) = False Then
            If MsgBox("�ٴ�����վҳ���е����廤�������IP��ַ���ò���ȷ�����ҽԺʹ�����ƶ����廤���������°滤ʿ����վ�޷�ʹ�����廤���ܡ�" & vbCrLf & "�������Ƿ�Ҫ������", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    ValidateData = True
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lst, Index, mrsPar)
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chk, Index, mrsPar)
End Sub

Private Sub optPASSVer_Click(Index As Integer)
    Dim strValue As String
    
    If Me.Visible Then
        strValue = IIF(optPASSVer(0).value, "3.0", "4.0")
        Call SetParChange(optPASSVer, Index, mrsPar, True, strValue)
        cmdSet.Visible = mobjPass.SetEnabled(cbo(cbo_������ҩ�ӿ�).ListIndex, strValue)
    End If
End Sub

Private Sub optPASSVer_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPASSVer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPASSVer, Index, mrsPar)
End Sub

Private Sub opt����Ŀ������_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt����Ŀ������, Index, mrsPar)
 
End Sub

Private Sub opt����Ŀ������_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt����Ŀ������_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt����Ŀ������, Index, mrsPar)
End Sub

Private Sub opt����Ŀ��סԺ_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt����Ŀ��סԺ, Index, mrsPar)
End Sub

Private Sub opt����Ŀ��סԺ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt����Ŀ��סԺ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt����Ŀ��סԺ, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)
    Dim strValue As String
    Dim strIP As String, strIPProt As String
    If Me.Visible Then
        Select Case Index
            Case txt_������ǰ�������
                Call SetParChange(opt�������, 1, mrsPar, True, IIF(opt�������(0).value, "0|0", IIF(opt�������(1).value, "1|" & NVL(txt(txt_������ǰ�������).Text, "0"), "2|" & NVL(txt(txt_������ǰ�������).Text, "0"))))
            Case txt_����������Ԥ��N��
                Call SetParChange(optEprRead, 2, mrsPar, True, Val(txt(txt_����������Ԥ��N��)))
                Call SetParChange(txt, Index, mrsPar, True, Val(txt(txt_����������Ԥ��N��)))
            Case txt_���廤��IP��ַ, txt_���廤��IP�˿�
                If txt(txt_���廤��IP��ַ).Text = "" Then
                    strValue = ""
                ElseIf txt(txt_���廤��IP�˿�).Text = "" Then
                    strValue = txt(txt_���廤��IP��ַ).Text
                Else
                    strValue = txt(txt_���廤��IP��ַ).Text & ":" & txt(txt_���廤��IP�˿�).Text
                End If
                Call SetParChange(txt, txt_���廤��IP��ַ, mrsPar, True, strValue) '���²���ֵ
                mrsPar.Filter = "ģ��=" & p�°�סԺ��ʿվ & " and ������='���廤��IP��ַ' "
                If Not mrsPar.EOF Then
                    If InStr(1, mrsPar!����ֵ & "", ":") <> 0 Then
                        strIP = Mid(mrsPar!����ֵ, 1, InStr(1, mrsPar!����ֵ, ":") - 1)
                        strIPProt = Mid(mrsPar!����ֵ, InStr(1, mrsPar!����ֵ, ":") + 1)
                    Else
                        strIP = mrsPar!����ֵ & ""
                        strIPProt = ""
                    End If
                    txt(txt_���廤��IP��ַ).ForeColor = IIF(txt(txt_���廤��IP��ַ).Text <> strIP, &HC0&, &H0&)
                    txt(txt_���廤��IP�˿�).ForeColor = IIF(txt(txt_���廤��IP�˿�).Text <> strIPProt, &HC0&, &H0&)
                End If
                cmdLink.Tag = ""
            Case Else
                Call SetParChange(txt, Index, mrsPar)
        End Select
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Select Case Index
    Case txt_������Һ�Զ�ˢ�²���, txt_������Һ����, txt_������Һ��ϵ��, txt_������Һ��Һ����, txt_������ҺƤ������, _
            txt_������ע����, txt_���¸��Ժϸ����, txt_���廤��IP��ַ, txt_���廤��IP�˿�
        Call zlControl.TxtSelAll(txt(Index))
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = Asc(gstrParSplit1) Or KeyAscii = Asc(gstrParSplit2) Then
        KeyAscii = 0
    End If
    
    Select Case Index
        Case txt_ǩ��ʹ��ͼƬ�߶�, txt_������N���Զ��۵�, txt_����������Ԥ��N��, txt_���Ӳ������ȱʡ����, _
            txt_���Ӳ�����������ˢ�¼��, txt_���Ӳ�������ȱʡ����, txt_���Ӳ��������������, _
            txt_������Һ�Զ�ˢ�²���, txt_������Һ����, txt_������Һ��ϵ��, txt_������Һ��Һ����, txt_������ҺƤ������, txt_������ע����, txt_���廤��IP�˿�
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        Case txt_���廤��IP��ַ
            If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = txt_������ǰ������� Then
        Call SetParTip(opt�������, 1, mrsPar)
    Else
        Call SetParTip(txt, Index, mrsPar)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case txt_������Һ�Զ�ˢ�²���
        If Val(txt(Index).Text) <> 0 And Val(txt(Index).Text) < 30 Then
            MsgBox "������Һ����ġ��Զ�ˢ�²����嵥��¼�벻��ȷ����Χ��0����ڵ���30����", vbInformation, gstrSysName
            txt(Index).Text = 30
        End If
    Case txt_������Һ����
        If Val(txt(Index).Text) < 10 Or Val(txt(Index).Text) > 100 Then
            MsgBox "������Һ����ġ�Ĭ�ϵ��١�¼�벻��ȷ����Χ��10-100����", vbInformation, gstrSysName
            txt(Index).Text = 40
        End If
    Case txt_������Һ��ϵ��
        If InStr(",10,15,20,", "," & Trim(txt(Index).Text) & ",") <= 0 Then
            MsgBox "������Һ����ġ�Ĭ�ϵ�ϵ����¼�벻��ȷ����Χ��10��15��20����", vbInformation, gstrSysName
            txt(Index).Text = 20
        End If
    Case txt_������Һ��Һ����
        If Val(txt(Index).Text) < 0 Or Val(txt(Index).Text) > 60 Then
            MsgBox "������Һ����ġ���Һ���ѡ�¼�벻��ȷ����Χ��0-60����", vbInformation, gstrSysName
            txt(Index).Text = 3
        End If
    Case txt_������ҺƤ������
        If Val(txt(Index).Text) < 0 Or Val(txt(Index).Text) > 60 Then
            MsgBox "������Һ����ġ�Ƥ�����ѡ�¼�벻��ȷ����Χ��0-60����", vbInformation, gstrSysName
            txt(Index).Text = 0
        End If
    Case txt_���廤��IP��ַ
        If txt(Index).Text = "" Then Exit Sub
        If IsIPAddress(txt(Index).Text) = False Then
            MsgBox "���廤�������IP��ַ¼�벻��ȷ�����飡", vbInformation, gstrSysName
            Cancel = True
        End If
    End Select
End Sub

Private Sub txtUD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtUD, Index, mrsPar)
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub



Private Function SetDeptInput(vsTmp As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByVal rsTmp As ADODB.Recordset) As Boolean
    '�ȼ���±�����Ƿ����
    Dim strTmp As String
    With vsTmp
        On Error Resume Next
        If vsTmp.Name = "vsStopDept" Then
            strTmp = mcolStop����("_" & rsTmp!ID)
        Else
            strTmp = mcol����("_" & rsTmp!ID)
        End If
        If Err.Number = 0 Then
            MsgBox "�ÿ����Ѿ����ڣ����������롣", vbInformation, gstrSysName
            .TextMatrix(lngRow, lngCol) = CStr(.Cell(flexcpData, lngRow, lngCol))
            Exit Function
        Else
            Err.Clear
        End If
        On Error GoTo 0
        
        If .TextMatrix(lngRow, lngCol + 4) <> "" Then
            If vsTmp.Name = "vsStopDept" Then
                Call mcolStop����.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
            Else
                Call mcol����.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
            End If
        End If
        
        .TextMatrix(lngRow, lngCol) = rsTmp!���� & ""
        .Cell(flexcpData, lngRow, lngCol) = rsTmp!���� & ""
        .TextMatrix(lngRow, lngCol + 4) = rsTmp!ID & ""
        If vsTmp.Name = "vsStopDept" Then
            Call mcolStop����.Add(rsTmp!ID & "", "_" & rsTmp!ID)
        Else
            Call mcol����.Add(rsTmp!ID & "", "_" & rsTmp!ID)
        End If
        SetDeptInput = True
    End With
End Function



Private Sub txtLocate_Change(Index As Integer)
    If Index = txt_Dept Then
        mlngPreFind = 1
    ElseIf Index = txt_Par Then
        txtLocate(Index).Tag = ""
    End If
End Sub

Private Sub txtLocate_GotFocus(Index As Integer)
    txtLocate(Index).SelStart = 0
    txtLocate(Index).SelLength = Len(txtLocate(Index).Text)
End Sub

Private Sub txtLocate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim strFind As String
        
        If Trim(txtLocate(Index).Text) = "" Then Exit Sub
        strFind = UCase(Trim(txtLocate(Index).Text))
        
        Select Case Index
        Case txt_Par
            Call LocatePar(txtLocate(Index), Me)
        Case txt_Dept
            If vsUnWriteDept.Visible Then
                Call LocateDept(strFind, vsUnWriteDept)
            End If
        End Select
    End If
End Sub

Private Sub LocateDept(ByVal strFind As String, ByRef objTmp As Object)
'���ܣ���鲻д�����Ŀ���
    Dim i As Long, j As Long
    Dim lngRows As Long, lngStart As Long
    Dim strCode As String, strName As String
    
    If TypeName(objTmp) = "ListBox" Then 'lst_�Զ�У�Բ���
        With objTmp
            lngRows = .ListCount - 1
            
            lngStart = IIF(mlngPreFind = 1, 0, mlngPreFind)
            For i = lngStart To .ListCount - 1
                strCode = Split(.List(i), "-")(0)
                strName = Split(.List(i), "-")(1)
                If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                    .ListIndex = i
                    .SetFocus
                    Exit For
                End If
            Next
        End With
        If i < lngRows Then
            mlngPreFind = i + 1
        Else
            If mlngPreFind = 1 Then
                MsgBox "û���ҵ�ƥ��ģ�������������ݡ�", vbInformation, Me.Caption
                txtLocate(txt_Dept).SetFocus
            Else
                MsgBox "ȫ�������ˣ�����û���ˡ�", vbInformation, Me.Caption
                mlngPreFind = 1
            End If
        End If
    Else
        '���ǵ��˹��ܵ�ʹ��Ƶ�ʵͣ���ʱ��֧����������
        With objTmp
            For i = 0 To .Rows - 1
                For j = 0 To .Cols - 1
                    If .ColHidden(j) = False Then
                        If .TextMatrix(i, j) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                            .Row = i: .Col = j
                            .ShowCell i, j
                            Exit Sub
                        End If
                    End If
                Next
            Next
            
            MsgBox "û���ҵ�ƥ��Ŀ��ң�������������ݡ�", vbInformation, Me.Caption
            txtLocate(txt_Dept).SetFocus
        End With
    End If
End Sub

Private Sub vsfDepartSign_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim l As Long
    Dim strTmp As String
    
    If vsfDepartSign.TextMatrix(Row, vsfDepartSign.ColIndex("ID")) = "" Then Exit Sub

    With vsfDepartSign
        For l = 1 To .Rows - 1
            If Val(.Cell(flexcpChecked, l, .ColIndex("����"))) <> Decode(Val(.RowData(l)), 1, 1, 2) Then
                vsfDepartSign.ForeColor = vbRed
                Exit Sub
            End If
        Next
        vsfDepartSign.ForeColor = vbBlack
    End With
End Sub

Private Sub vsfDepartSign_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDepartSign
        If Row = .Rows - 1 Then
            If .TextMatrix(Row, .ColIndex("ID")) = "" And .ColKey(Col) = "����" Then Cancel = 1
            vsfDepartSign.TextMatrix(Row, Col) = ""
        Else
            If .ColKey(Col) <> "����" Then Cancel = 1
        End If
    End With
End Sub

Private Sub vsfDepartSign_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    cmdDepartSelect.Visible = False
End Sub

Private Sub vsfDepartSign_DblClick()
    If vsfDepartSign.TextMatrix(vsfDepartSign.RowSel, vsfDepartSign.ColIndex("ID")) = "" Then Exit Sub
    vsfDepartSign.TextMatrix(vsfDepartSign.RowSel, vsfDepartSign.ColIndex("����")) = IIF(vsfDepartSign.TextMatrix(vsfDepartSign.RowSel, vsfDepartSign.ColIndex("����")) = "1", "0", "1")
    Call vsfDepartSign_AfterEdit(vsfDepartSign.RowSel, vsfDepartSign.ColSel)
End Sub

Private Sub vsfDepartSign_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    Dim lngCol As Long, lngRow As Long, X As Long, Y As Long, rsTmp As ADODB.Recordset, blnCancel As Boolean
    Dim vPoint As POINTAPI, i As Long, blnChange As Boolean, strNewValue As String
    
    vPoint = zlControl.GetCoordPos(vsfDepartSign.hwnd, cmdDepartSelect.Left, cmdDepartSelect.Top)
    lngCol = vsfDepartSign.Col: lngRow = vsfDepartSign.Row
    
    gstrSQL = "Select b.Id, b.����, b.����,b.���롡" & vbNewLine & _
                    "from (Select ����id, ����id, ����ֵ From Zldeptparas Where ����id = (Select ID From zlParameters Where ������ = 'ǩ��ʹ��ͼƬ')) A,���ű� B" & vbNewLine & _
                    "Where b.Id = a.����id(+) And a.����id Is Null And (Instr(b.����, Upper([1])) > 0 Or Instr(b.����, Upper([1])) > 0)" & vbNewLine & _
                    "And (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                    "Order By b.����"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "", False, "����", "��ѡ����", False, False, True, vPoint.X, vPoint.Y, 225, blnCancel, True, True, vsfDepartSign.EditText)
        
    If blnCancel Then
        vsfDepartSign.TextMatrix(lngRow, lngCol) = ""
    Else
        If rsTmp Is Nothing Then
            vsfDepartSign.TextMatrix(lngRow, lngCol) = ""
            Exit Sub
        End If
        For i = 1 To vsfDepartSign.Rows - 1
            If vsfDepartSign.TextMatrix(i, vsfDepartSign.ColIndex("ID")) = rsTmp!ID Then
                vsfDepartSign.TextMatrix(lngRow, lngCol) = ""
                Exit Sub
            End If
        Next
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("ID")) = rsTmp!ID & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("����")) = rsTmp!���� & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("����")) = rsTmp!���� & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("����")) = "1"
        blnChange = True
    End If

    If blnChange Then
        If vsfDepartSign.TextMatrix(vsfDepartSign.Rows - 1, vsfDepartSign.ColIndex("ID")) <> "" Then
            vsfDepartSign.Rows = vsfDepartSign.Rows + 1
        End If

        With vsfDepartSign
            For i = 1 To .Rows - 1
                If Val(.Cell(flexcpChecked, i, .ColIndex("����"))) <> Decode(Val(.RowData(i)), 1, 1, 2) Then
                    vsfDepartSign.ForeColor = vbRed
                    Exit Sub
                End If
            Next
            vsfDepartSign.ForeColor = vbBlack
        End With
        vsfDepartSign.Row = vsfDepartSign.Rows - 1
        Call vsfDepartSign.ShowCell(vsfDepartSign.Rows - 1, 0)
    End If
End Sub

Private Sub vsfDepartSign_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long, lngRowHeight As Long, lngTop As Long, lngLeft As Long, i As Integer, blnScroll As Boolean, lngRowsHeight As Long

    lngRow = vsfDepartSign.MouseRow: lngCol = vsfDepartSign.MouseCol
    If lngRow = vsfDepartSign.Rows - 1 And lngCol >= 0 Then
        For i = 1 To vsfDepartSign.Rows - 1
            lngRowsHeight = lngRowsHeight + vsfDepartSign.RowHeight(i) + IIF(vsfDepartSign.RowHeight(i) = vsfDepartSign.RowHeight(0), 30, 0)
        Next
        If lngRowsHeight > vsfDepartSign.ClientHeight - vsfDepartSign.RowHeight(0) - 30 Then blnScroll = True '��Ϊ����������CMD��λ����Ӧ�����ƶ�240
            
        For i = vsfDepartSign.TopRow To lngRow - 1
            lngTop = lngTop + vsfDepartSign.RowHeight(i) + IIF(vsfDepartSign.RowHeight(i) = vsfDepartSign.RowHeight(0), 30, 0)
        Next
        lngTop = lngTop + vsfDepartSign.RowHeight(0) + 30
        'lngLeft = IIF(lngCol = 1, 1600, vsfDepartSign.Width - 10 - IIF(blnScroll, 240, 0)) - cmdDepartSelect.Width - 30
        lngLeft = vsfDepartSign.Width - 10 - IIF(blnScroll, 240, 0) - cmdDepartSelect.Width - 30
        lngRowHeight = vsfDepartSign.RowHeight(lngRow)
        Call cmdDepartSelect.Move(lngLeft, lngTop, cmdDepartSelect.Width, lngRowHeight)
        cmdDepartSelect.Visible = True
    Else
        cmdDepartSelect.Visible = False
    End If
        
    Call SetParTip(vsfDepartSign, 0, mrsPar)
End Sub

Private Sub vsfDrugStore_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strObjTag As String
    On Error Resume Next
    strObjTag = GetDrugTag(Index, vsfDrugStore(Index).MouseRow, vsfDrugStore(Index).MouseCol)
    Call SetParTip(vsfDrugStore, Index, mrsPar, strObjTag)
End Sub

Private Function GetDrugTag(ByVal lngIndex As Long, ByVal lngRow As Long, ByVal lngCol As Long) As String
    With vsfDrugStore(lngIndex)
        If lngCol = .ColIndex("����") Then
            GetDrugTag = IIF(lngIndex = cbo_����ҩ�����Ҷ��շ���, "����", "סԺ") & "����" & .TextMatrix(lngRow, .ColIndex("���"))
        ElseIf lngCol = .ColIndex("ȱʡ") Then
            GetDrugTag = IIF(lngIndex = cbo_����ҩ�����Ҷ��շ���, "����", "סԺ") & "ȱʡ" & .TextMatrix(lngRow, .ColIndex("���"))
        ElseIf lngCol = .ColIndex("��ҩ����") Then
            GetDrugTag = .TextMatrix(lngRow, .ColIndex("���")) & "����"
        End If
    End With
End Function

Private Sub vsfDrugStore_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDrugStore(Index).ColIndex("����") Then
        Call Set����ҩ��(Index, Row, True)
    ElseIf Col = vsfDrugStore(Index).ColIndex("����") Then
        Call Setȱʡҩ��(Index)
    End If
    If Col <> vsfDrugStore(Index).ColIndex("��ҩ����") Then Cancel = True
End Sub

Private Sub vsfDrugStore_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore(Index)
        Select Case Col
        Case .ColIndex("����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("ȱʡ")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("��ҩ����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick(Index As Integer)
    With vsfDrugStore(Index)
        If .MouseCol = .ColIndex("ȱʡ") Then
            Call Setȱʡҩ��(Index)
        ElseIf .MouseCol = .ColIndex("ҩ��") Then
            Call Set����ҩ��(Index, .Row, True)
        ElseIf .MouseCol = .ColIndex("����") And .MouseRow = .FixedRows - 1 Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                Call Set����ҩ��(Index, i)
            Next
        End If
    End With
End Sub

Private Sub vsfDrugStore_EnterCell(Index As Integer)

    With vsfDrugStore(Index)
        If .Row > 0 Then
            .FocusRect = flexFocusLight
        End If
    End With
End Sub

Private Sub vsfDrugStore_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If vsfDrugStore(Index).Col = vsfDrugStore(Index).ColIndex("ȱʡ") Then
            Call Setȱʡҩ��(Index)
        End If
    End If
End Sub

Private Sub SetRecordҩ��(ByVal lngIndex As Long, ByVal lngRow As Long, ByVal lngCol As Long)
    Dim strTmp As String, i As Long
    Dim strValue As String
    
    strTmp = Replace(Replace(GetDrugTag(lngIndex, lngRow, lngCol), "����", ""), "סԺ", "")
    
    With vsfDrugStore(lngIndex)
        If strTmp Like "����*" Then
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("���")) = .TextMatrix(lngRow, .ColIndex("���")) Then
                    If .TextMatrix(i, .ColIndex("����")) <> 0 Then
                        strValue = strValue & "," & .RowData(i)
                    End If
                End If
            Next
            strValue = Mid(strValue, 2)
        ElseIf strTmp Like "ȱʡ*" Then
            strValue = IIF(.TextMatrix(lngRow, .ColIndex("ȱʡ")) = "��", .RowData(lngRow), "")
        ElseIf strTmp Like "*����" Then
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("���")) = .TextMatrix(lngRow, .ColIndex("���")) Then
                    If .TextMatrix(i, .ColIndex("��ҩ����")) <> "�Զ�����" Then
                        strValue = strValue & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("��ҩ����"))
                    End If
                End If
            Next
            strValue = Mid(strValue, 2)
        End If
    End With
    IIF(lngIndex = cbo_����ҩ�����Ҷ��շ���, mrs����ҩ������, mrsסԺҩ������).Fields(strTmp).value = strValue
    SST.Tag = "���޸�"
End Sub

Private Sub Setȱʡҩ��(ByVal Index As Integer)
'���ܣ����õ�ǰ�е�ȱʡҩ����ͬʱ������ͬ���͵������е�ȱʡҩ��
    Dim i As Long
    
    With vsfDrugStore(Index)
        If Val("" & .Cell(flexcpData, .Row, .ColIndex("ȱʡ"))) = 0 Then  '�ò��������޸ĵ������
            If .TextMatrix(.Row, .ColIndex("ȱʡ")) = "��" Then
                .TextMatrix(.Row, .ColIndex("ȱʡ")) = ""
            Else
                '��û����Ȩ���޸Ŀ���ʱ�ҿ���Ϊ0��false)ʱ����������ȱʡ
                If Not (Val(.TextMatrix(.Row, .ColIndex("����"))) = 0 And Val("" & .Cell(flexcpData, .Row, .ColIndex("����"))) = 1) Then
                    'ͬ����������ȡ��ȱʡ
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(.Row, .ColIndex("���")) = .TextMatrix(i, .ColIndex("���")) Then
                            If .TextMatrix(i, .ColIndex("ȱʡ")) = "��" Then .TextMatrix(i, .ColIndex("ȱʡ")) = ""
                        End If
                    Next
                    .TextMatrix(.Row, .ColIndex("����")) = -1    '�Զ�����Ϊ����
                    Call SetRecordҩ��(Index, .Row, .ColIndex("����"))
                    .TextMatrix(.Row, .ColIndex("ȱʡ")) = "��"
                Else
                    MsgBox "���õ�ǰҩ��Ϊȱʡʱ����ͬʱ����ǰҩ������Ϊ���ã�" & vbNewLine & "��û���޸Ŀ���ҩ����Ȩ�ޡ�", vbInformation, gstrSysName
                End If
            End If
            Call SetRecordҩ��(Index, .Row, .ColIndex("ȱʡ"))
        Else
            MsgBox "��û���޸�ȱʡҩ����Ȩ�ޡ�", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub Set����ҩ��(ByVal Index As Integer, ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = False)
'���ܣ����õ�ǰ�еĿ���ҩ����ͬʱ����ǰ�е�ȱʡҩ��

    With vsfDrugStore(Index)
        If Val("" & .Cell(flexcpData, lngRow, .ColIndex("����"))) = 0 Then   '�ò��������޸ĵ������
            If Val(.TextMatrix(lngRow, .ColIndex("����"))) = -1 Then
                '��ǰ���ҹ�ѡ����
                If Not (Val("" & .Cell(flexcpData, lngRow, .ColIndex("ȱʡ"))) = 1 And .TextMatrix(lngRow, .ColIndex("ȱʡ")) = "��") Then
                    .TextMatrix(lngRow, .ColIndex("����")) = 0
                    .TextMatrix(lngRow, .ColIndex("ȱʡ")) = ""
                    Call SetRecordҩ��(Index, lngRow, .ColIndex("ȱʡ"))
                    Call SetRecordҩ��(Index, lngRow, .ColIndex("����"))
                Else
                    If blnAsk Then
                        MsgBox "ȡ����ǰҩ������ʱ����ͬʱȡ����ǰҩ��ȱʡ��" & vbNewLine & "��û���޸�ȱʡҩ����Ȩ�ޡ�", vbInformation, gstrSysName
                    End If
                End If
            Else
                .TextMatrix(lngRow, .ColIndex("����")) = -1    '�Զ�����Ϊ����
                Call SetRecordҩ��(Index, lngRow, .ColIndex("����"))
            End If
        Else
            If blnAsk Then
                MsgBox "��û���޸Ŀ���ҩ����Ȩ�ޡ�", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub


Private Sub vsfDrugStore_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Index = cbo_����ҩ�����Ҷ��շ��� Then
        If Col = vsfDrugStore(Index).ColIndex("��ҩ����") Then
            Call SetRecordҩ��(Index, Row, Col)
        End If
    End If
End Sub

Private Sub vsfEpr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim l As Long, strNewValue As String
    If Not (Index = 1 Or Index = 2) Then Exit Sub '0����������Ҫ 1 2����顢����ָ�����Ҽ���Ա��Χ
    
    If KeyCode = vbKeyDelete Then
        For l = vsfEpr(Index).Row To vsfEpr(Index).Rows - 1
            If l <> vsfEpr(Index).Rows - 1 Then
                vsfEpr(Index).TextMatrix(l, 0) = vsfEpr(Index).TextMatrix(l + 1, 0)
                vsfEpr(Index).TextMatrix(l, 1) = vsfEpr(Index).TextMatrix(l + 1, 1)
                vsfEpr(Index).TextMatrix(l, 2) = vsfEpr(Index).TextMatrix(l + 1, 2)
                vsfEpr(Index).TextMatrix(l, 3) = vsfEpr(Index).TextMatrix(l + 1, 3)
            Else
                vsfEpr(Index).TextMatrix(l, 0) = ""
                vsfEpr(Index).TextMatrix(l, 1) = ""
                vsfEpr(Index).TextMatrix(l, 2) = ""
                vsfEpr(Index).TextMatrix(l, 3) = ""
            End If
        Next
        vsfEpr(Index).Rows = vsfEpr(Index).Rows - 1
        If vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 0) <> "" And vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 2) <> "" Then
            vsfEpr(Index).Rows = vsfEpr(Index).Rows + 1
        End If
        vsfEpr(Index).RowHeight(vsfEpr(Index).Rows - 1) = 240
        vsfEpr(Index).AutoSize 3, 3
        vsfEpr(Index).Row = vsfEpr(Index).Rows - 1
        
        For l = 1 To vsfEpr(Index).Rows - 1
            If vsfEpr(Index).TextMatrix(l, 0) <> "" And vsfEpr(Index).TextMatrix(l, 2) <> "" Then
                strNewValue = strNewValue & ";" & vsfEpr(Index).TextMatrix(l, 0) & "," & vsfEpr(Index).TextMatrix(l, 2)
             End If
        Next
        If strNewValue <> "" Then
            strNewValue = Mid(strNewValue, 2)
        End If
        
        Call SetParChange(vsfEpr, Index, mrsPar, True, strNewValue)
    End If
End Sub

Private Sub vsfEpr_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If Not (Index = 1 Or Index = 2) Then Exit Sub '0����������Ҫ 1 2����顢����ָ�����Ҽ���Ա��Χ

If KeyAscii = vbKeyReturn Then
Dim lngCol As Long, lngRow As Long, X As Long, Y As Long, rsTmp As New ADODB.Recordset, blnCancel As Boolean
Dim vPoint As POINTAPI, l As Long, blnChange As Boolean, strNewValue As String, strOldValue As String
    vPoint = zlControl.GetCoordPos(vsfEpr(Index).hwnd, cmdEprSelect(Index).Left, cmdEprSelect(Index).Top)
    lngCol = vsfEpr(Index).Col: lngRow = vsfEpr(Index).Row
    strOldValue = vsfEpr(Index).TextMatrix(lngRow, lngCol)
    If lngCol = 1 Then '����
        gstrSQL = "Select Distinct a.���,a.Id, a.����, a.����, c.���� As ����" & vbNewLine & _
                "From ��Ա�� A, ��Ա����˵�� B, ���ű� C, ������Ա D" & vbNewLine & _
                "Where a.Id = b.��Աid And c.Id = d.����id And d.��Աid = a.Id And d.ȱʡ = 1 And" & vbNewLine & _
                "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And b.��Ա���� In ('ҽ��') And" & vbNewLine & _
                "      (Instr(a.����,[1])>0 or Instr(a.����,Upper([1]))>0 or Instr(a.���,[1])>0 ) " & vbNewLine & _
                "Order By a.���"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "", False, "����", "��ѡ��һ�������Ա", False, False, True, vPoint.X, vPoint.Y, 225, blnCancel, True, True, vsfEpr(Index).EditText)
        If blnCancel Then
            vsfEpr(Index).TextMatrix(lngRow, 1) = strOldValue
        Else
            If rsTmp.State = adStateClosed Then
                vsfEpr(Index).TextMatrix(lngRow, 1) = strOldValue: vsfEpr(Index).EditText = strOldValue
            ElseIf rsTmp.EOF Then
                vsfEpr(Index).TextMatrix(lngRow, 1) = strOldValue: vsfEpr(Index).EditText = strOldValue
            Else
                blnChange = True
                vsfEpr(Index).TextMatrix(lngRow, 0) = rsTmp!ID
                vsfEpr(Index).TextMatrix(lngRow, 1) = rsTmp!����
            End If
        End If
    ElseIf lngCol = 3 Then '����
        gstrSQL = "Select a.Id,a.����, a.����, a.����" & vbNewLine & _
                    "From ���ű� A, ��������˵�� B" & vbNewLine & _
                    "Where a.Id = b.����id And b.�������� In ('�ٴ�') And b.������� In (2, 3) And" & vbNewLine & _
                    "      (To_Char(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' Or a.����ʱ�� Is Null) And" & vbNewLine & _
                    "      (Instr(a.����, [1])>0 or Instr(a.����,Upper([1]))>0 or Instr(a.����,[1])>0  ) " & vbNewLine & _
                    "Order By a.����"
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, gstrSQL, 0, "", False, "", "��ѡ��һ���������˿���", False, False, True, vPoint.X, vPoint.Y, 225, blnCancel, True, True, vsfEpr(Index).EditText)
        If blnCancel Then
            vsfEpr(Index).TextMatrix(lngRow, 3) = strOldValue
        Else
            If rsTmp.State = adStateClosed Then
                vsfEpr(Index).TextMatrix(lngRow, 3) = strOldValue: vsfEpr(Index).EditText = strOldValue
            ElseIf rsTmp.EOF Then
                vsfEpr(Index).TextMatrix(lngRow, 3) = strOldValue: vsfEpr(Index).EditText = strOldValue
            Else
                blnChange = True
                Do Until rsTmp.EOF
                    If rsTmp.AbsolutePosition = 1 Then
                        vsfEpr(Index).TextMatrix(lngRow, 2) = rsTmp!ID
                        vsfEpr(Index).TextMatrix(lngRow, 3) = rsTmp!����
                    Else
                        vsfEpr(Index).TextMatrix(lngRow, 2) = vsfEpr(Index).TextMatrix(lngRow, 2) & "," & rsTmp!ID
                        vsfEpr(Index).TextMatrix(lngRow, 3) = vsfEpr(Index).TextMatrix(lngRow, 3) & vbCrLf & rsTmp!����
                    End If
                    rsTmp.MoveNext
                Loop
            End If
            vsfEpr(Index).AutoSize 3, 3
        End If
    End If
    
    If blnChange Then
        If vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 0) <> "" And vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 2) <> "" Then
            vsfEpr(Index).Rows = vsfEpr(Index).Rows + 1
        End If
        
        For l = 1 To vsfEpr(Index).Rows - 1
            If vsfEpr(Index).TextMatrix(l, 0) <> "" And vsfEpr(Index).TextMatrix(l, 2) <> "" Then
                strNewValue = strNewValue & ";" & vsfEpr(Index).TextMatrix(l, 0) & "," & vsfEpr(Index).TextMatrix(l, 2)
             End If
        Next
        If strNewValue <> "" Then
            strNewValue = Mid(strNewValue, 2)
        End If
        
        Call SetParChange(vsfEpr, Index, mrsPar, True, strNewValue)
    End If
End If
End Sub

Private Sub vsfEpr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngRow As Long, lngCol As Long, lngRowHeight As Long, lngTop As Long, lngLeft As Long, i As Integer, blnScroll As Boolean, lngRowsHeight As Long
    If Index = 1 Or Index = 2 Then
        lngRow = vsfEpr(Index).MouseRow: lngCol = vsfEpr(Index).MouseCol
        If lngRow >= 1 And lngCol >= 0 Then
            For i = 1 To vsfEpr(Index).Rows - 1
                lngRowsHeight = lngRowsHeight + vsfEpr(Index).RowHeight(i) + IIF(vsfEpr(Index).RowHeight(i) = vsfEpr(Index).RowHeight(0), 30, 0)
            Next
            If lngRowsHeight > vsfEpr(Index).ClientHeight - vsfEpr(Index).RowHeight(0) - 30 Then blnScroll = True '��Ϊ����������CMD��λ����Ӧ�����ƶ�240
            
            For i = vsfEpr(Index).TopRow To lngRow - 1
                lngTop = lngTop + vsfEpr(Index).RowHeight(i) + IIF(vsfEpr(Index).RowHeight(i) = vsfEpr(Index).RowHeight(0), 30, 0)
            Next
            lngTop = lngTop + vsfEpr(Index).RowHeight(0) + 30
            lngLeft = IIF(lngCol = 1, 1600, vsfEpr(Index).Width - 10 - IIF(blnScroll, 240, 0)) - cmdEprSelect(Index).Width - 30
            lngRowHeight = vsfEpr(Index).RowHeight(lngRow)
            Call cmdEprSelect(Index).Move(lngLeft, lngTop, cmdEprSelect(Index).Width, lngRowHeight)
            cmdEprSelect(Index).Visible = True
        End If
    End If
    
    Call SetParTip(vsfEpr, Index, mrsPar)
End Sub

Private Sub vsfMecItem_DblClick()
    Call cmdModify_Click
End Sub

Private Sub vsfMecItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsfMecItem.Row > 0 Then
            Call cmdDelete_Click
        End If
    End If
End Sub

Private Sub vsfWaittingMixDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim l As Long
    Dim strTmp As String
    
    With vsfWaittingMixDept
        For l = 1 To .Rows - 1
            If Val(.TextMatrix(l, .ColIndex("����"))) <> 0 Then
                strTmp = strTmp & "," & Trim(.TextMatrix(l, .ColIndex("ID")))
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        
        Call SetParChange(vsfWaittingMixDept, 0, mrsPar, True, strTmp)
    End With
End Sub

Private Sub vsfWaittingMixDept_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = vsfWaittingMixDept.ColKey(Col) <> "����"
End Sub

Private Sub vsfWaittingMixDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsfWaittingMixDept, 0, mrsPar)
End Sub

Private Sub vsUnWriteDept_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    
    If Me.Visible Then
        strValue = Get����(vsUnWriteDept)
        Call SetParChange(vsUnWriteDept, 0, mrsPar, True, strValue)
    End If
End Sub

Private Sub vsStopDept_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    
    If Me.Visible Then
        strValue = Get����(vsStopDept)
        Call SetParChange(vsStopDept, 0, mrsPar, True, strValue)
    End If
End Sub

Private Sub vsUnWriteDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    
    With vsUnWriteDept
        If KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsUnWriteDept_KeyPress(KeyCode)
        ElseIf KeyCode = vbKeyDelete Then
            Call mcol����.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
            .TextMatrix(.Row, .Col) = ""
            .Cell(flexcpData, .Row, .Col) = ""
            .TextMatrix(.Row, .Col + 4) = ""
            
            Call vsUnWriteDept_AfterEdit(.Row, .Col)
        End If
        If KeyCode = vbKeyReturn Then
            If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Call EnterNextCell(vsUnWriteDept)
        End If
    End With
End Sub

Private Sub vsStopDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    
    With vsStopDept
        If KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsStopDept_KeyPress(KeyCode)
        ElseIf KeyCode = vbKeyDelete Then
            Call mcolStop����.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
            .TextMatrix(.Row, .Col) = ""
            .Cell(flexcpData, .Row, .Col) = ""
            .TextMatrix(.Row, .Col + 4) = ""
            
            Call vsStopDept_AfterEdit(.Row, .Col)
        End If
        If KeyCode = vbKeyReturn Then
            If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Call EnterNextCell(vsStopDept)
        End If
    End With
End Sub


Private Sub vsUnWriteDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsUnWriteDept.TextMatrix(Row, Col + 4) = "" Then vsUnWriteDept.TextMatrix(Row, Col) = ""
End Sub

Private Sub vsStopDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsStopDept.TextMatrix(Row, Col + 4) = "" Then vsStopDept.TextMatrix(Row, Col) = ""
End Sub


Private Sub vsUnWriteDept_KeyPress(KeyAscii As Integer)
    If vsUnWriteDept.Editable = flexEDNone Then Exit Sub

    With vsUnWriteDept
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsUnWriteDept_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsStopDept_KeyPress(KeyAscii As Integer)
    If vsStopDept.Editable = flexEDNone Then Exit Sub

    With vsStopDept
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsStopDept_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub Load����(vsTmp As VSFlexGrid, ByVal strIn As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    vsTmp.Clear
    If strIn = "" Then Exit Sub
    
    strIn = Replace(strIn, "|", ",")
    strSQL = "select id,���� from ���ű� where id in (Select Column_Value From Table(f_Num2list([1]))) Order by ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIn)
    If rsTmp.EOF Then Exit Sub
    
    With vsTmp
        lngRow = (rsTmp.RecordCount + 3) \ 4
        If lngRow > 5 Then .Rows = lngRow
        
        For i = 1 To rsTmp.RecordCount
            If vsTmp.Name = "vsStopDept" Then
                Call mcolStop����.Add(rsTmp!ID & "", "_" & rsTmp!ID)
            Else
                Call mcol����.Add(rsTmp!ID & "", "_" & rsTmp!ID)
            End If
            lngRow = (i - 1) \ 4
            lngCol = (i - 1) Mod 4
            
            .TextMatrix(lngRow, lngCol) = rsTmp!����
            .Cell(flexcpData, lngRow, lngCol) = rsTmp!���� & ""
            .TextMatrix(lngRow, lngCol + 4) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Function Get����(vsTmp As VSFlexGrid) As String
    Dim i As Integer
    Dim j As Integer
    Dim strIds As String
    
    With vsTmp
        For i = 0 To .Rows - 1
            For j = 0 To 3
                If .TextMatrix(i, j) <> "" Then
                    strIds = strIds & "|" & Val(.TextMatrix(i, j + 4))
                End If
            Next
        Next
    End With
    
    Get���� = Mid(strIds, 2)
End Function



Private Sub vsUnWriteDept_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsUnWriteDept
    If .Editable = flexEDNone Then
        .FocusRect = flexFocusLight
        .ComboList = ""
    Else
        .FocusRect = flexFocusSolid
        .ComboList = "..."
    End If
    End With
End Sub

Private Sub vsStopDept_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsStopDept
    If .Editable = flexEDNone Then
        .FocusRect = flexFocusLight
        .ComboList = ""
    Else
        .FocusRect = flexFocusSolid
        .ComboList = "..."
    End If
    End With
End Sub

Private Sub vsUnWriteDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsUnWriteDept
    If .Editable = flexEDNone Then
        .FocusRect = flexFocusLight
        .ComboList = ""
    Else
        .FocusRect = flexFocusSolid
        .ComboList = "..."
    End If
    End With
End Sub

Private Sub vsStopDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsStopDept
    If .Editable = flexEDNone Then
        .FocusRect = flexFocusLight
        .ComboList = ""
    Else
        .FocusRect = flexFocusSolid
        .ComboList = "..."
    End If
    End With
End Sub


Private Sub vsUnWriteDept_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean, blnInput As Boolean
    
    On Error GoTo errH
    strSQL = "select Distinct a.id,a.����,a.����,a.���� from ���ű� a,��������˵�� b where a.id=b.����id" & _
        " and b.��������='�ٴ�' And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) Order by A.����"
    With vsUnWriteDept
        vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "�ٴ�����", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            blnInput = SetDeptInput(vsUnWriteDept, Row, Col, rsTmp)
            If blnInput Then
                Call EnterNextCell(vsUnWriteDept)
                Call vsUnWriteDept_AfterEdit(Row, Col)
            Else
                Call vsUnWriteDept_AfterRowColChange(-1, -1, Row, Col)
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п��õĿ������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsStopDept_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean, blnInput As Boolean
    
    On Error GoTo errH
    strSQL = "select Distinct a.id,a.����,a.����,a.���� from ���ű� a,��������˵�� b where a.id=b.����id" & _
        " and b.��������='�ٴ�' And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) Order by A.����"
    With vsStopDept
        vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "�ٴ�����", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            blnInput = SetDeptInput(vsStopDept, Row, Col, rsTmp)
            If blnInput Then
                Call EnterNextCell(vsStopDept)
                Call vsStopDept_AfterEdit(Row, Col)
            Else
                Call vsStopDept_AfterRowColChange(-1, -1, Row, Col)
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п��õĿ������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsUnWriteDept_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean, blnInput As Boolean
    
    With vsUnWriteDept
        If .EditText = CStr(.Cell(flexcpData, Row, Col)) Then
            Call EnterNextCell(vsUnWriteDept)
            Exit Sub
        End If
        strInput = UCase(.EditText)
        strSQL = "select Distinct a.id,a.����,a.����,a.���� from ���ű� a,��������˵�� b where a.id=b.����id" & _
            " and b.��������='�ٴ�' And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
            " Order by A.����"
        With vsUnWriteDept
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�ٴ�����", False, "", "", False, False, True, _
                vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
            If Not rsTmp Is Nothing Then
                If SetDeptInput(vsUnWriteDept, Row, Col, rsTmp) Then
                    .EditText = .TextMatrix(Row, Col)
                    Call EnterNextCell(vsUnWriteDept)
                    Exit Sub
                End If
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��Ŀ��ҡ�", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
        End With
        Call vsUnWriteDept_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
    End With
End Sub

Private Sub vsStopDept_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean, blnInput As Boolean
    
    With vsStopDept
        If .EditText = CStr(.Cell(flexcpData, Row, Col)) Then
            Call EnterNextCell(vsStopDept)
            Exit Sub
        End If
        strInput = UCase(.EditText)
        strSQL = "select Distinct a.id,a.����,a.����,a.���� from ���ű� a,��������˵�� b where a.id=b.����id" & _
            " and b.��������='�ٴ�' And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
            " Order by A.����"
        With vsStopDept
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�ٴ�����", False, "", "", False, False, True, _
                vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
            If Not rsTmp Is Nothing Then
                If SetDeptInput(vsStopDept, Row, Col, rsTmp) Then
                    .EditText = .TextMatrix(Row, Col)
                    Call EnterNextCell(vsStopDept)
                    Exit Sub
                End If
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��Ŀ��ҡ�", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
        End With
        Call vsStopDept_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
    End With
End Sub


Private Sub SaveDepartSign()
'���沿��ʹ��ͼƬǩ������
    Dim i As Long
    
    On Error GoTo ErrHandle
    With vsfDepartSign
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpChecked, i, .ColIndex("����"))) <> Decode(Val(.RowData(i)), 1, flexChecked, flexUnchecked) Then
                Call zlDatabase.SetPara("ǩ��ʹ��ͼƬ", IIF(.TextMatrix(i, .ColIndex("����")) = "0", "0", "1"), glngSys, p�����ڲ�����, , Val(.TextMatrix(i, .ColIndex("ID"))))
            End If
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Saveҽ������()
'����ҽ�����ݶ���
    Dim blnTrans As Boolean

    On Error GoTo ErrHandle
    If cmdAdvice.Tag = "���޸�" Then
        
        gcnOracle.BeginTrans: blnTrans = True
        gstrSQL = "zl_ҽ�����ݶ���_Delete"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        mrsAdvice.Filter = 0
        Do While Not mrsAdvice.EOF
            If Not IsNull(mrsAdvice!ҽ������) Then
                gstrSQL = "zl_ҽ�����ݶ���_Insert('" & mrsAdvice!������� & "','" & Replace(mrsAdvice!ҽ������, "'", "''") & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
            mrsAdvice.MoveNext
        Loop
        gcnOracle.CommitTrans: blnTrans = False
        cmdAdvice.Tag = ""
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Saveҩ�����Ҷ���()
'����ҩ�����Ҷ���
    Dim i As Long, strOutFilter As String, strInFilter
    Dim str���� As String
    
    If SST.Tag <> "���޸�" Then
        mrsPar.Filter = "������='ҩ�����Ҷ��շ���' and �޸�״̬=1 "
        If Not mrsPar.EOF Then SST.Tag = "���޸�"
    End If
    
    If SST.Tag = "���޸�" Then
        Call zlDatabase.DelDeptPara("����ȱʡ��ҩ��", glngSys, p����ҽ���´�)
        Call zlDatabase.DelDeptPara("����ȱʡ��ҩ��", glngSys, p����ҽ���´�)
        Call zlDatabase.DelDeptPara("����ȱʡ��ҩ��", glngSys, p����ҽ���´�)
        Call zlDatabase.DelDeptPara("���������ҩ��", glngSys, p����ҽ���´�)
        Call zlDatabase.DelDeptPara("������ó�ҩ��", glngSys, p����ҽ���´�)
        Call zlDatabase.DelDeptPara("���������ҩ��", glngSys, p����ҽ���´�)
        Call zlDatabase.DelDeptPara("����ȱʡ���ϲ���", glngSys, p����ҽ���´�)
        Call zlDatabase.DelDeptPara("������÷��ϲ���", glngSys, p����ҽ���´�)
        strOutFilter = mrs����ҩ������.Filter: strInFilter = mrsסԺҩ������.Filter
        mrs����ҩ������.Filter = 0: mrsסԺҩ������.Filter = 0
        If mrs����ҩ������.RecordCount > 0 Then mrs����ҩ������.MoveFirst
        If mrsסԺҩ������.RecordCount > 0 Then mrsסԺҩ������.MoveFirst
        Do While Not mrs����ҩ������.EOF
            If mrs����ҩ������!����IDs & "" <> "" Then
                str���� = str���� & ";" & mrs����ҩ������!����IDs
                For i = 0 To UBound(Split(mrs����ҩ������!����IDs, ","))
                    Call zlDatabase.SetPara("����ȱʡ��ҩ��", mrs����ҩ������!ȱʡ��ҩ�� & "", glngSys, p����ҽ���´�, , Split(mrs����ҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("����ȱʡ��ҩ��", mrs����ҩ������!ȱʡ��ҩ�� & "", glngSys, p����ҽ���´�, , Split(mrs����ҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("����ȱʡ��ҩ��", mrs����ҩ������!ȱʡ��ҩ�� & "", glngSys, p����ҽ���´�, , Split(mrs����ҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("���������ҩ��", mrs����ҩ������!������ҩ�� & "", glngSys, p����ҽ���´�, , Split(mrs����ҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("������ó�ҩ��", mrs����ҩ������!���ó�ҩ�� & "", glngSys, p����ҽ���´�, , Split(mrs����ҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("���������ҩ��", mrs����ҩ������!������ҩ�� & "", glngSys, p����ҽ���´�, , Split(mrs����ҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("����ȱʡ���ϲ���", mrs����ҩ������!ȱʡ���ϲ��� & "", glngSys, p����ҽ���´�, , Split(mrs����ҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("������÷��ϲ���", mrs����ҩ������!���÷��ϲ��� & "", glngSys, p����ҽ���´�, , Split(mrs����ҩ������!����IDs, ",")(i))
                Next
            End If
            mrs����ҩ������.MoveNext
        Loop
        Call zlDatabase.SetPara("ҩ�����Ҷ��շ���", Mid(str����, 2), glngSys, p����ҽ���´�)
        str���� = ""
        Call zlDatabase.DelDeptPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�)
        Call zlDatabase.DelDeptPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�)
        Call zlDatabase.DelDeptPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�)
        Call zlDatabase.DelDeptPara("סԺ������ҩ��", glngSys, pסԺҽ���´�)
        Call zlDatabase.DelDeptPara("סԺ���ó�ҩ��", glngSys, pסԺҽ���´�)
        Call zlDatabase.DelDeptPara("סԺ������ҩ��", glngSys, pסԺҽ���´�)
        Call zlDatabase.DelDeptPara("סԺȱʡ���ϲ���", glngSys, pסԺҽ���´�)
        Call zlDatabase.DelDeptPara("סԺ���÷��ϲ���", glngSys, pסԺҽ���´�)
        Do While Not mrsסԺҩ������.EOF
            If mrsסԺҩ������!����IDs & "" <> "" Then
                str���� = str���� & ";" & mrsסԺҩ������!����IDs
                For i = 0 To UBound(Split(mrsסԺҩ������!����IDs, ","))
                    Call zlDatabase.SetPara("סԺȱʡ��ҩ��", mrsסԺҩ������!ȱʡ��ҩ�� & "", glngSys, pסԺҽ���´�, , Split(mrsסԺҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("סԺȱʡ��ҩ��", mrsסԺҩ������!ȱʡ��ҩ�� & "", glngSys, pסԺҽ���´�, , Split(mrsסԺҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("סԺȱʡ��ҩ��", mrsסԺҩ������!ȱʡ��ҩ�� & "", glngSys, pסԺҽ���´�, , Split(mrsסԺҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("סԺ������ҩ��", mrsסԺҩ������!������ҩ�� & "", glngSys, pסԺҽ���´�, , Split(mrsסԺҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("סԺ���ó�ҩ��", mrsסԺҩ������!���ó�ҩ�� & "", glngSys, pסԺҽ���´�, , Split(mrsסԺҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("סԺ������ҩ��", mrsסԺҩ������!������ҩ�� & "", glngSys, pסԺҽ���´�, , Split(mrsסԺҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("סԺȱʡ���ϲ���", mrsסԺҩ������!ȱʡ���ϲ��� & "", glngSys, pסԺҽ���´�, , Split(mrsסԺҩ������!����IDs, ",")(i))
                    Call zlDatabase.SetPara("סԺ���÷��ϲ���", mrsסԺҩ������!���÷��ϲ��� & "", glngSys, pסԺҽ���´�, , Split(mrsסԺҩ������!����IDs, ",")(i))
                Next
            End If
            mrsסԺҩ������.MoveNext
        Loop
        Call zlDatabase.SetPara("ҩ�����Ҷ��շ���", Mid(str����, 2), glngSys, pסԺҽ���´�)
        
        mrs����ҩ������.Filter = IIF(strOutFilter = "0", 0, strOutFilter): mrsסԺҩ������.Filter = IIF(strInFilter = "0", 0, strInFilter)
        SST.Tag = ""
    End If
End Sub

Private Sub Setҩ�����Ҷ��ղ�������(ByVal lngPar As Long)
'���ܣ��ı��������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strPar As String
    Dim blnTrans As Boolean
    Dim i As Long
    
    On Error GoTo errH
    If lngPar = 0 Then
        strPar = ",0,0,0,'" & gstrUserName & "','�޸Ĺ�������(����ҩ�����հ�������������)Ӱ��',1)"
    ElseIf lngPar = 1 Then
        strPar = ",0,1,0,'" & gstrUserName & "','�޸Ĺ�������(����ҩ�����հ�������������)Ӱ��',0)"
    End If
    If strPar <> "" Then
        strSQL = "Select n.id From zlParameters N Where n.ģ�� In (1252, 1253)" & vbNewLine & _
                "And  n.������ In ('����ȱʡ��ҩ��', '����ȱʡ��ҩ��', '����ȱʡ��ҩ��', '����ȱʡ���ϲ���', '���������ҩ��', '������ó�ҩ��', '���������ҩ��', 'סԺȱʡ��ҩ��', 'סԺȱʡ��ҩ��'," & vbNewLine & _
                "                'סԺȱʡ��ҩ��', 'סԺȱʡ���ϲ���', 'סԺ������ҩ��', 'סԺ���ó�ҩ��', 'סԺ������ҩ��','������÷��ϲ���','סԺ���÷��ϲ���')"

        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To rsTmp.RecordCount
            zlDatabase.ExecuteProcedure "Zl_Parameters_Change(" & rsTmp!ID & strPar, Me.Caption
            rsTmp.MoveNext
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    If Me.Visible Then
        If Index = lst_����ִ���Զ����ҽ����� Then
            Call SetValueסԺִ��
            Frame14.Tag = "���޸�"
        Else
            Call SetParChange(lst, Index, mrsPar)
        End If
    End If
End Sub

Private Sub txtUD_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtUD(Index).Text) > ud(Index).Max Or Val(txtUD(Index).Text) < ud(Index).Min Then
        txtUD(Index).Text = ud(Index).value
    End If
End Sub

Private Sub txtUD_Change(Index As Integer)
    Dim strValue As String
    
    If Me.Visible Then
        Select Case Index
        Case txtud_����ҹ�࿪ʼʱ��, txtud_����ҹ�����ʱ��
                    strValue = txtUD(txtud_����ҹ�࿪ʼʱ��).Text & ";" & txtUD(txtud_����ҹ�����ʱ��).Text
                    If Index = txtud_����ҹ�࿪ʼʱ�� Then
                        Call SetParChange(txtUD, txtud_����ҹ�����ʱ��, mrsPar, True, strValue)
                    Else
                        Call SetParChange(txtUD, txtud_����ҹ�࿪ʼʱ��, mrsPar, True, strValue)
                    End If
                    Call SetParChange(txtUD, Index, mrsPar, True, strValue)
        Case Else
                Call SetParChange(txtUD, Index, mrsPar)
        End Select
    End If
End Sub

Private Sub txtUD_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtUD(Index))
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub cmdAdvice_Click()
    If frmAdviceDefine.ShowMe(Me, mrsAdvice) Then
        '���Ϊ�ѱ仯,��Ҫ����
        cmdAdvice.Tag = "���޸�"
    End If
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cbo_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    Dim arrIndex As Variant, i As Integer
    
    Select Case Index
    Case cbo_������ҩ�ӿ�
        '����ʱ�ɼ�
        lblPassVer.Visible = InStr(",1,2,", "," & cbo(Index).ListIndex & ",") > 0
        optPASSVer(0).Visible = InStr(",1,2,", "," & cbo(Index).ListIndex & ",") > 0
        optPASSVer(1).Visible = InStr(",1,2,", "," & cbo(Index).ListIndex & ",") > 0
        
        If cbo(Index).ListIndex = 0 Then    'δ���ýӿ�
            chk(chk_����ҩ��).Enabled = False
            chk(chk_����ҩ��).value = 0
            chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Enabled = False
            chk(chk_��ֹ�´ﳬ����ҩƷҽ��).value = 0
            chk(chk_����Ժ��ִ�н���ҩƷ).Enabled = False
            chk(chk_����Ժ��ִ�н���ҩƷ).value = 0
            chk(chk_����ҩƷҪ����дԭ��).Enabled = False
            chk(chk_����ҩƷҪ����дԭ��).value = 0
        
            chk(chk_�ӿڵ�����־_��ͨ).Visible = False  '��ͨʱ�ɼ�
            chk(chk_ʹ��ϵͳ����_����).Visible = False  '����ʱ�ɼ�
      
            '̫Ԫͨʱ�ɼ�
            cbo(cmd_����������Դ).Visible = False
            lblInfo(lbl_����������Դ).Visible = False
            cmdSet.Visible = False
        Else
            chk(chk_����ҩ��).Enabled = True
            chk(chk_����Ժ��ִ�н���ҩƷ).Enabled = chk(chk_����ҩ��).value = 0
            chk(chk_����ҩƷҪ����дԭ��).Enabled = chk(chk_����ҩ��).value = 1 And InStr(",1,2,3,6,", "," & cbo(cbo_������ҩ�ӿ�).ListIndex & ",") > 0
            If Not chk(chk_����ҩƷҪ����дԭ��).Enabled Then chk(chk_����ҩƷҪ����дԭ��).value = 0
            
            If cbo(Index).ListIndex = 1 Then  '����
                chk(chk_ʹ��ϵͳ����_����).Visible = True
                chk(chk_ʹ��ϵͳ����_����).Enabled = True
                optPASSVer(0).Caption = "����3.0"
                optPASSVer(1).Caption = "����4.0"
            Else
                chk(chk_ʹ��ϵͳ����_����).Visible = False
                chk(chk_ʹ��ϵͳ����_����).Enabled = False
            End If

            If cbo(Index).ListIndex = 2 Then  '��ͨ
                chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Enabled = True
                chk(chk_�ӿڵ�����־_��ͨ).Visible = True
                optPASSVer(0).Caption = "CS��"
                optPASSVer(1).Caption = "BS��"
            Else
                chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Enabled = False
                chk(chk_��ֹ�´ﳬ����ҩƷҽ��).value = 0
                chk(chk_�ӿڵ�����־_��ͨ).Visible = False
            End If
            If cbo(Index).ListIndex = 3 Then    '̫Ԫͨ
                cbo(cmd_����������Դ).ListIndex = 0
                cbo(cmd_����������Դ).Visible = True
                lblInfo(lbl_����������Դ).Visible = True
                cbo(cmd_����������Դ).Enabled = True
                lblInfo(lbl_����������Դ).Enabled = True
            Else
                cbo(cmd_����������Դ).Visible = False
                lblInfo(lbl_����������Դ).Visible = False
            End If
            If mobjPass Is Nothing Then
                On Error Resume Next
                Set mobjPass = CreateObject("zlPassInterface.clsPass")
                Err.Clear: On Error GoTo 0
            End If
            If Not mobjPass Is Nothing Then
                If optPASSVer(1).Visible Then strValue = IIF(optPASSVer(1).value, "4.0", "3.0")
                cmdSet.Visible = mobjPass.SetEnabled(cbo(Index).ListIndex, strValue)
            End If
        End If
    Case cbo_��ҩ�䷽
        blnValue = True
        strValue = IIF(cbo(cbo_��ҩ�䷽).ListIndex = 1, 4, 3)
    Case cbo_����ҩ�����Ҷ��շ���, cbo_סԺҩ�����Ҷ��շ���
        Call Set����ҩ������(Index)
    Case cbo_סԺ����ִ���Զ���ɷ���
        Call Set�Զ���ɷ���(Index)
    Case cbo_��������˳��, cbo_��¶�½�˳��
        blnValue = True
        strValue = cbo(cbo_��������˳��).ListIndex & ";" & cbo(cbo_��¶�½�˳��).ListIndex
        If Index = cbo_��������˳�� Then
            If Me.Visible Then Call SetParChange(cbo, cbo_��¶�½�˳��, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_��������˳��, mrsPar, blnValue, strValue)
        End If
    Case cbo_���������쳣��, cbo_��¶�½��쳣��
        blnValue = True
        strValue = cbo(cbo_���������쳣��).ListIndex & ";" & cbo(cbo_��¶�½��쳣��).ListIndex
        If Index = cbo_���������쳣�� Then
            If Me.Visible Then Call SetParChange(cbo, cbo_��¶�½��쳣��, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_���������쳣��, mrsPar, blnValue, strValue)
        End If
    Case cbo_������־����, cbo_������־λ��
        blnValue = True
        strValue = cbo(cbo_������־����).ListIndex & ";" & cbo(cbo_������־λ��).ListIndex
        If Index = cbo_������־���� Then
            If Me.Visible Then Call SetParChange(cbo, cbo_������־λ��, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_������־����, mrsPar, blnValue, strValue)
        End If
    Case cbo_��������ʾ, cbo_�쳣����ʾ
        blnValue = True
        strValue = cbo(cbo_��������ʾ).ListIndex & ";" & cbo(cbo_�쳣����ʾ).ListIndex
        If Index = cbo_��������ʾ Then
            If Me.Visible Then Call SetParChange(cbo, cbo_�쳣����ʾ, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_��������ʾ, mrsPar, blnValue, strValue)
        End If
    Case cbo_��Ժ�Զ���־, cbo_����Զ���־, cbo_ת���Զ���־, cbo_�����Զ���־, cbo_�����Զ���־, cbo_��Ժ�Զ���־, _
        cbo_�����Զ���־, cbo_�����Զ���־, cbo_�����Զ���־, cbo_ת�����Զ���־
        blnValue = True
        strValue = cbo(cbo_��Ժ�Զ���־).ListIndex & ";" & cbo(cbo_����Զ���־).ListIndex & ";" & cbo(cbo_ת���Զ���־).ListIndex & ";" & _
            cbo(cbo_�����Զ���־).ListIndex & ";" & cbo(cbo_�����Զ���־).ListIndex & ";" & cbo(cbo_��Ժ�Զ���־).ListIndex & ";" & _
            cbo(cbo_�����Զ���־).ListIndex & ";" & cbo(cbo_�����Զ���־).ListIndex & ";" & cbo(cbo_�����Զ���־).ListIndex & ";" & _
            cbo(cbo_ת�����Զ���־).ListIndex
        arrIndex = Array(cbo_��Ժ�Զ���־, cbo_����Զ���־, cbo_ת���Զ���־, cbo_�����Զ���־, cbo_�����Զ���־, cbo_��Ժ�Զ���־, _
            cbo_�����Զ���־, cbo_�����Զ���־, cbo_�����Զ���־, cbo_ת�����Զ���־)
        For i = 0 To UBound(arrIndex)
            If Index <> CInt(arrIndex(i)) Then
                If Me.Visible Then Call SetParChange(cbo, CInt(arrIndex(i)), mrsPar, blnValue, strValue)
            End If
        Next i
    Case cbo_��Ѫ�ɼ�Ĭ����������
        blnValue = True
        If cbo(cbo_��Ѫ�ɼ�Ĭ����������).ListIndex >= 0 Then
            strValue = zlCommFun.GetNeedName(cbo(cbo_��Ѫ�ɼ�Ĭ����������).List(cbo(cbo_��Ѫ�ɼ�Ĭ����������).ListIndex), "-")
        End If
    Case cbo_С��ȱʡ��ʶ
        blnValue = True
        strValue = cbo(cbo_С��ȱʡ��ʶ).ListIndex & ";" & chk(chk_С��ȱʡ��ʶ).value
        If Me.Visible Then Call SetParChange(cbo, cbo_С��ȱʡ��ʶ, mrsPar, blnValue, strValue)
    End Select
    
    If Me.Visible Then
        Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
    End If
End Sub

Private Sub Set����ҩ������(ByVal lngIndex As Long)
    Dim strSQL As String, rsTmp As Recordset
    Dim strDeptIDs As String
    Dim i As Long, j As Long
    Dim strDefault As String, strDSIDs As String, str���� As String, strPar As String
    
    If lngIndex = cbo_����ҩ�����Ҷ��շ��� Then
        mrs����ҩ������.Filter = "����=" & cbo(lngIndex).ItemData(cbo(lngIndex).ListIndex)
        If mrs����ҩ������.RecordCount > 0 Then strDeptIDs = mrs����ҩ������!����IDs & ""
    ElseIf lngIndex = cbo_סԺҩ�����Ҷ��շ��� Then
        mrsסԺҩ������.Filter = "����=" & cbo(lngIndex).ItemData(cbo(lngIndex).ListIndex)
        If mrsסԺҩ������.RecordCount > 0 Then strDeptIDs = mrsסԺҩ������!����IDs & ""
    End If
    strSQL = "select ID,���� From ���ű� Where ID in(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDeptIDs)
    vsUseDept(lngIndex).Enabled = True
    vsfDrugStore(lngIndex).Enabled = True
 
    vsUseDept(lngIndex).Rows = 0
    Do While Not rsTmp.EOF
        vsUseDept(lngIndex).Rows = vsUseDept(lngIndex).Rows + 1
        vsUseDept(lngIndex).TextMatrix(vsUseDept(lngIndex).Rows - 1, 0) = rsTmp!���� & ""
        vsUseDept(lngIndex).Cell(flexcpData, vsUseDept(lngIndex).Rows - 1, 0) = rsTmp!ID & ""

        rsTmp.MoveNext
    Loop
    If vsUseDept(lngIndex).Rows = 0 Then
        vsUseDept(lngIndex).Rows = 1
    Else
        '����ҩ������
        With vsfDrugStore(lngIndex)
            For i = 1 To .Rows - 1
                strDefault = IIF(lngIndex = cbo_����ҩ�����Ҷ��շ���, mrs����ҩ������, mrsסԺҩ������).Fields("ȱʡ" & .TextMatrix(i, .ColIndex("���"))).value & ""
                strDSIDs = "," & IIF(lngIndex = cbo_����ҩ�����Ҷ��շ���, mrs����ҩ������, mrsסԺҩ������).Fields("����" & .TextMatrix(i, .ColIndex("���"))).value & ","
                
                If Val(.RowData(i)) = Val(strDefault) Then
                    .TextMatrix(i, .ColIndex("ȱʡ")) = "��"
                    .TextMatrix(i, .ColIndex("����")) = -1   'true
                Else
                    .TextMatrix(i, .ColIndex("ȱʡ")) = ""
                    .TextMatrix(i, .ColIndex("����")) = IIF(InStr(strDSIDs, "," & Val(.RowData(i)) & ",") > 0, -1, 0)
                End If
                strPar = IIF(lngIndex = cbo_����ҩ�����Ҷ��շ���, mrs����ҩ������, mrsסԺҩ������).Fields("ȱʡ���ϲ���").value & ""
                 
            Next
        End With
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    Dim strIndexs As String '���ڲ����Ƿֳɰ�λ��ȡ��Ҫ���⴦��
    Dim varTmp As Variant
    Dim i As Long
    
    Select Case Index
        Case chk_�����Ǽ���Ч����
            txtUD(ud_�����Ǽ���Ч����).Enabled = chk(Index).value = 1
            txtUD(ud_�����Ǽ���Ч����).BackColor = IIF(chk(Index).value = 1, RGB(255, 255, 255), Me.BackColor)
            ud(ud_�����Ǽ���Ч����).Enabled = txtUD(ud_�����Ǽ���Ч����).Enabled
            strValue = IIF(chk(Index).value = 1, ud(ud_�����Ǽ���Ч����).value, "0")
            blnValue = True
        Case chk_���ﴦ����������
            txtUD(ud_���ﴦ����������).Enabled = chk(Index).value = 1
            txtUD(ud_���ﴦ����������).BackColor = IIF(chk(Index).value = 1, RGB(255, 255, 255), Me.BackColor)
            ud(ud_���ﴦ����������).Enabled = txtUD(ud_���ﴦ����������).Enabled
            strValue = IIF(chk(Index).value = 1, ud(ud_���ﴦ����������).value, "0")
            blnValue = True
        Case chk_����ҩ��ּ�����
            chk(chk_����ҩ��ʹ���Ա�ҩ).Enabled = chk(Index).value = 1
            chk(chk_��ҽ��С����п���ҩ�����).Enabled = chk(Index).value = 1
        Case chk_����ҩ��
            chk(chk_����Ժ��ִ�н���ҩƷ).value = 0
            chk(chk_����ҩƷҪ����дԭ��).value = 0
            chk(chk_����Ժ��ִ�н���ҩƷ).Enabled = chk(Index).value = 0
            chk(chk_����ҩƷҪ����дԭ��).Enabled = chk(Index).value = 1 And InStr(",1,2,3,6,", "," & cbo(cbo_������ҩ�ӿ�).ListIndex & ",") > 0
        Case chk_�����ּ�����
            If chk(Index).value = 0 Then
                chk(chk_������Ȩ����).value = 0
                chk(chk_�����ּ����).value = 0
                chk(chk_����ҽʦ�ﵽ�����ȼ��������).value = 0
            End If
            chk(chk_������Ȩ����).Enabled = chk(Index).value = 1
            chk(chk_�����ּ����).Enabled = chk(Index).value = 1
            chk(chk_����ҽʦ�ﵽ�����ȼ��������).Enabled = chk(Index).value = 1
        Case chk_��Ѫ�ּ�����
            If chk(Index).value = 0 Then
                chk(chk_��Ѫ�����������).value = 0
                chk(chk_��Ѫ���������м�������ҽʦ).value = 0
            End If
            chk(chk_��Ѫ�����������).Enabled = chk(Index).value = 1
            chk(chk_��Ѫ���������м�������ҽʦ).Enabled = chk(Index).value = 1
        Case chk_����Ѫ�����ϵͳ
            If mblnUseBlood Then
                If chk(Index).value = 0 Then
                    chk(chk_��Ѫҽ�����ͺ���ܷ�Ѫ).value = 0
                    chk(chk_��Ѫ���벻��ʾѪҺ���).value = 0
                    chk(chk_�´���Ѫ����ʱȷ����Ѫ��Ϣ).value = 0
                    chk(chk_ѪҺ���պ������ִ�еǼ�).value = 0
                    If cbo(cbo_��Ѫ�ɼ�Ĭ����������).ListCount > 0 Then
                        cbo(cbo_��Ѫ�ɼ�Ĭ����������).ListIndex = 0
                    End If
                End If
                chk(chk_��Ѫҽ�����ͺ���ܷ�Ѫ).Enabled = chk(Index).value = 1
                chk(chk_��Ѫ���벻��ʾѪҺ���).Enabled = chk(Index).value = 1
                chk(chk_�´���Ѫ����ʱȷ����Ѫ��Ϣ).Enabled = chk(Index).value = 1
                chk(chk_ѪҺ���պ������ִ�еǼ�).Enabled = chk(Index).value = 1
                lblBloodManager.Enabled = chk(Index).value = 1
                cbo(cbo_��Ѫ�ɼ�Ĭ����������).Enabled = chk(Index).value = 1
            End If
        Case chk_ҽ��ִ����Ч����
            txtUD(ud_ҽ��ִ����Ч����).Enabled = chk(Index).value = 1
            txtUD(ud_ҽ��ִ����Ч����).BackColor = IIF(chk(Index).value = 1, RGB(255, 255, 255), Me.BackColor)
            ud(ud_ҽ��ִ����Ч����).Enabled = txtUD(ud_ҽ��ִ����Ч����).Enabled
            strValue = IIF(chk(Index).value = 1, ud(ud_ҽ��ִ����Ч����).value, "999")
            blnValue = True
        Case chk_ҽ������ʱ��������ԭ��
            Call SetVsfEditable(vsUnWriteDept, chk(Index).value = 1)
        Case chk_ͣ��ʱ¼��ԭ��
            Call SetVsfEditable(vsStopDept, chk(Index).value = 1)
        Case chk_�����ڷ�ҩ���ͽ���ʱ��
            dtp�ڷ�����ʱ��.Enabled = chk(Index).value = 1
            blnValue = True
            strValue = IIF(chk(Index).value = 1, "1|" & dtp�ڷ�����ʱ��.value, "0")
        Case chk_��Ѫҽ��ִ�к���Ҫ�˶�, chk_Ƥ��ҽ��ִ�к���Ҫ�˶�
            blnValue = True
            strValue = chk(chk_��Ѫҽ��ִ�к���Ҫ�˶�).value & chk(chk_Ƥ��ҽ��ִ�к���Ҫ�˶�).value
            strIndexs = chk_��Ѫҽ��ִ�к���Ҫ�˶� & "|" & chk_Ƥ��ҽ��ִ�к���Ҫ�˶�
        Case chk_סԺ�����Զ�ִ�г���, chk_סԺ�����Զ�ִ������
            blnValue = True
            strValue = chk(chk_סԺ�����Զ�ִ�г���).value & chk(chk_סԺ�����Զ�ִ������).value
            lst(lst_����ִ���Զ����ҽ�����).Enabled = Val(strValue) <> 0
        Case chk_סԺ�´��Զ�����
            cmdAdviceSortSet.Enabled = chk(Index).value = 1
        Case chk_סԺ�´��Ƥ��
            optδƤ������ҽ��(0).Enabled = chk(Index).value = 1
            optδƤ������ҽ��(1).Enabled = optδƤ������ҽ��(0).Enabled
        Case chk_·��ִ�л���ҽ������, chk_·��ִ�л��ڻ�ʿ����
            blnValue = True
            strValue = chk(chk_·��ִ�л���ҽ������).value & chk(chk_·��ִ�л��ڻ�ʿ����).value
        Case chk_����·��ִ�л���
            If chk(Index).value = 0 Then
                chk(chk_·��ִ�л���ҽ������).value = 0
                chk(chk_·��ִ�л���ҽ������).Enabled = False
                chk(chk_·��ִ�л��ڻ�ʿ����).value = 0
                chk(chk_·��ִ�л��ڻ�ʿ����).Enabled = False
            Else
                chk(chk_·��ִ�л���ҽ������).Enabled = True
                chk(chk_·��ִ�л��ڻ�ʿ����).Enabled = True
            End If
        Case chk_����ǰһ�첻���������ɽ���·����Ŀ
            If chk(Index).value = 0 Then
                chk(chk_������ǰ���������·����Ŀ).value = 0
                chk(chk_������ǰ���������·����Ŀ).Enabled = False
            Else
                chk(chk_������ǰ���������·����Ŀ).Enabled = True
            End If
        Case chk_������ת�ƻ�ҳ
            chk(chk_ת�ƻ�ҳ�������д�ӡ�ؿ�ҽ��).Enabled = chk(Index).value = 1
        Case chk_ǩ��ʹ��ͼƬ
            'chk(chk_ǩ��ʹ��ԭͼ).Enabled = (chk(chk_ǩ��ʹ��ͼƬ).value = 1)
            txt(txt_ǩ��ʹ��ͼƬ�߶�).Enabled = (chk(chk_ǩ��ʹ��ͼƬ).value = 1)
        Case chk_ǩ��ʹ��ԭͼ
            txt(txt_ǩ��ʹ��ͼƬ�߶�).Enabled = Not (chk(chk_ǩ��ʹ��ԭͼ).value = 1)
        Case chk_����ͼ����ʾ������
            cbo(cbo_��������ʾ).Enabled = chk(Index).value = 1
            cbo(cbo_�쳣����ʾ).Enabled = chk(Index).value = 1
        Case chk_���뵥���û���������, chk_���뵥���û���סԺ���, chk_���뵥���û����������, chk_���뵥���û���סԺ����, chk_���뵥���û���������Ѫ, _
                chk_���뵥���û���סԺ��Ѫ, chk_���뵥���û�����������, chk_���뵥���û���סԺ����, chk_���뵥���û��ڻ���
            blnValue = True
            strValue = Get���뵥���û���
            strIndexs = chk_���뵥���û��������� & "|" & chk_���뵥���û���סԺ��� & "|" & chk_���뵥���û���������� & "|" & chk_���뵥���û���סԺ���� & "|" & chk_���뵥���û���������Ѫ & "|" & _
                    chk_���뵥���û���סԺ��Ѫ & "|" & chk_���뵥���û����������� & "|" & chk_���뵥���û���סԺ���� & "|" & chk_���뵥���û��ڻ���
        Case chk_�������뵥�����ʹ�����뵥�´�ҽ������, chk_�������뵥�����ʹ�����뵥�´�ҽ��סԺ
            blnValue = True
            strValue = chk(chk_�������뵥�����ʹ�����뵥�´�ҽ������).value & chk(chk_�������뵥�����ʹ�����뵥�´�ҽ��סԺ).value
            strIndexs = chk_�������뵥�����ʹ�����뵥�´�ҽ������ & "|" & chk_�������뵥�����ʹ�����뵥�´�ҽ��סԺ
        Case chk_����ҩ�����հ�������������
            SST.Enabled = chk(Index).value = 0
        Case chk_��Ӧ��ݻ����ļ�
            chk(chk_��������ͬ��).Enabled = chk(chk_��Ӧ��ݻ����ļ�).value = 1
        Case chk_ҩƷҽ����ͬ���಻��·����ҽ��
            If chk(chk_ҩƷҽ����ͬ���಻��·����ҽ��).value = 1 Then chk(chk_ҩƷҽ����ƥ��Ϊ·������Ŀ).value = 0
        Case chk_ҩƷҽ����ƥ��Ϊ·������Ŀ
            If chk(chk_ҩƷҽ����ƥ��Ϊ·������Ŀ).value = 1 Then chk(chk_ҩƷҽ����ͬ���಻��·����ҽ��).value = 0
        Case chk_С��ȱʡ��ʶ
            blnValue = True
            strValue = cbo(cbo_С��ȱʡ��ʶ).ListIndex & ";" & chk(chk_С��ȱʡ��ʶ).value
    End Select
    
    If Me.Visible Then
        If strIndexs = "" Then
            Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
        Else
            varTmp = Split(strIndexs, "|")
            For i = 0 To UBound(varTmp)
                Call SetParChange(chk, Val(varTmp(i)), mrsPar, blnValue, strValue)
            Next
        End If
    End If
End Sub

Private Sub cmd���ͻ������_Click(Index As Integer)
    Dim j As Long
    
    If SendPriceType.Tab = 0 Then
        j = lst_���﷢�ͻ������
    Else
        j = lst_סԺ���ͻ������
    End If
    Call SetLstSelected(lst(j), Index = 0)
End Sub


Private Sub cmd���﷢��һ�ŵ������_Click(Index As Integer)
    If lst(lst_���﷢��һ�ŵ������).Enabled = False Then Exit Sub
    Call SetLstSelected(lst(lst_���﷢��һ�ŵ������), Index = 0)
End Sub

Private Sub cmd���﷢�ͼ�����_Click(Index As Integer)
    Call SetLstSelected(lst(lst_���﷢�ͼ�����), Index = 0)
End Sub

Private Sub cmdסԺ�����Ժ���_Click(Index As Integer)
    Call SetLstSelected(lst(lst_סԺ�����Ժ���), Index = 0)
End Sub


Private Sub opt���ڷ����ջ�_Click(Index As Integer)
     '��������ʱ����ʹ���Զ�������뵥
     
    chk(chk_�����ջ��Զ���˱���).Enabled = (Index = 1)
    If Index = 0 Then chk(chk_�����ջ��Զ���˱���).value = 0
    
    If Me.Visible Then Call SetParChange(opt���ڷ����ջ�, Index, mrsPar)
End Sub

Private Sub opt���͵��ݹ���_Click(Index As Integer)
    lst(lst_���﷢��һ�ŵ������).Enabled = opt���͵��ݹ���(0).value
    
    chk(chk_һ������һ�ŵ���).Enabled = opt���͵��ݹ���(0).value
    chk(chk_�������ҽ������ʱһ����鷢��Ϊһ�ŵ���).Enabled = opt���͵��ݹ���(0).value
    
    If Me.Visible Then Call SetParChange(opt���͵��ݹ���, Index, mrsPar)
End Sub

Private Sub opt���͵�������_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt���͵�������, Index, mrsPar)
End Sub


Private Sub opt��Ѫ���뵥��ӡ_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt��Ѫ���뵥��ӡ, Index, mrsPar)
End Sub

Private Sub optδƤ������ҽ��_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optδƤ������ҽ��, Index, mrsPar)
End Sub

Private Sub optסԺҽ������ӡ_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optסԺҽ������ӡ, Index, mrsPar)
End Sub

Private Sub dtp�ڷ�����ʱ��_Change()
    If Me.Visible Then
        Call SetParChange(dtp�ڷ�����ʱ��, 0, mrsPar, True, "1|" & dtp�ڷ�����ʱ��.value)
    End If
End Sub


Private Sub vsUnWriteDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsUnWriteDept, 0, mrsPar)
End Sub

Private Sub vsStopDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsStopDept, 0, mrsPar)
End Sub

Private Sub opt���ڷ����ջ�_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt���ڷ����ջ�, Index, mrsPar)
End Sub

Private Sub opt���͵��ݹ���_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt���͵��ݹ���, Index, mrsPar)
End Sub

Private Sub opt���͵�������_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt���͵�������, Index, mrsPar)
End Sub


Private Sub opt��Ѫ���뵥��ӡ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt��Ѫ���뵥��ӡ, Index, mrsPar)
End Sub

Private Sub optδƤ������ҽ��_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optδƤ������ҽ��, Index, mrsPar)
End Sub


Private Sub optסԺҽ������ӡ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optסԺҽ������ӡ, Index, mrsPar)
End Sub

Private Sub cmdAdviceSortSet_Click()
    frmParAdviceSort.mbytFun = 1
    frmParAdviceSort.Show vbModal, Me
End Sub

Private Sub vsUseDept_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsUseDept(Index).Editable = flexEDNone Then
        vsUseDept(Index).FocusRect = flexFocusLight
        vsUseDept(Index).ComboList = ""
    Else
        vsUseDept(Index).FocusRect = flexFocusSolid
        vsUseDept(Index).ComboList = "..."
    End If
End Sub


Private Sub vsUseDept_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim vPoint As POINTAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset, i As Long, strIds As String
    Dim strTmp As String, blnCheck As Boolean
    Dim strWhere As String
    
    If Col = 0 Then
        With vsUseDept(Index)
            For i = 0 To .Rows - 1
                If .Cell(flexcpData, i, 0) <> "" Then
                    strIds = strIds & "," & .Cell(flexcpData, i, 0)
                End If
            Next
            strIds = Mid(strIds, 2)
            
            If Index = cbo_����ҩ�����Ҷ��շ��� Then
                strWhere = " and �������� in ('�ٴ�','���','����','����','����','Ӫ��') And B.������� in (1,3)"
            ElseIf Index = cbo_סԺҩ�����Ҷ��շ��� Then
                strWhere = " and �������� in ('�ٴ�','���','����','����','����','Ӫ��') And B.������� in (2,3)"
            ElseIf Index = cbo_סԺ����ִ���Զ���ɷ��� Then
                strWhere = " And (  b.�������� = '�ٴ�' And ((b.������� In (2, 3)) Or (b.������� = 1 And Exists (Select 1 From ��λ״����¼ C Where b.����id = c.����id)))" & vbNewLine & _
                            "    Or b.�������� = '����' And b.������� In (1, 2, 3))"
            End If
                        
            strSQL = "select distinct ID,Decode(instr([1],',' || ID || ','),0,0,1) AS CHECKID,����,����" & _
                " from ���ű� A,��������˵�� B" & _
                " where A.ID=B.����ID " & strWhere & _
                "       and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " order by ����"
            
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "����", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, "," & strIds & ",")
                    
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    .Rows = 0
                    Do While Not rsTmp.EOF
                        If Index = cbo_����ҩ�����Ҷ��շ��� Then
                            blnCheck = Check����ҩ������(mrs����ҩ������, Val(mrs����ҩ������!���� & ""), Val(rsTmp!ID & ""), rsTmp!���� & "")
                        ElseIf Index = cbo_סԺҩ�����Ҷ��շ��� Then
                            blnCheck = Check����ҩ������(mrsסԺҩ������, Val(mrsסԺҩ������!���� & ""), Val(rsTmp!ID & ""), rsTmp!���� & "")
                        ElseIf Index = cbo_סԺ����ִ���Զ���ɷ��� Then
                            blnCheck = Check����ҩ������(mrsסԺִ�ж���, Val(mrsסԺִ�ж���!���� & ""), Val(rsTmp!ID & ""), rsTmp!���� & "")
                        End If
                        If blnCheck Then
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, Col) = rsTmp!���� & ""
                            .Cell(flexcpData, .Rows - 1, Col) = rsTmp!ID & ""
                            strTmp = strTmp & "," & .Cell(flexcpData, .Rows - 1, Col)
                        End If
                        rsTmp.MoveNext
                    Loop
                    If Index = cbo_����ҩ�����Ҷ��շ��� Then
                        mrs����ҩ������!����IDs = Mid(strTmp, 2)
                        mrs����ҩ������.Update
                        Call SetParChange(vsUseDept, Index, mrsPar, True, Get����ҩ������(mrs����ҩ������))
                    ElseIf Index = cbo_סԺҩ�����Ҷ��շ��� Then
                        mrsסԺҩ������!����IDs = Mid(strTmp, 2)
                        mrsסԺҩ������.Update
                        Call SetParChange(vsUseDept, Index, mrsPar, True, Get����ҩ������(mrsסԺҩ������))
                    ElseIf Index = cbo_סԺ����ִ���Զ���ɷ��� Then
                        mrsסԺִ�ж���!����IDs = Mid(strTmp, 2)
                        mrsסԺִ�ж���.Update
                        Call SetParChange(vsUseDept, Index, mrsPar, True, Get����ҩ������(mrsסԺִ�ж���))
                    End If
                Else
                    MsgBox "��ǰû��ѡ��Ŀ��ҡ�", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Function Check����ҩ������(ByVal rsTmp As Recordset, ByVal lng��ǰ���� As Long, ByVal lng����ID As Long, ByVal str�������� As String) As Boolean
    Dim strValue As String, strFilter As String
    Dim blnYes As Boolean
    
    If rsTmp.RecordCount = 0 Then Exit Function
    strFilter = rsTmp.Filter
    rsTmp.Filter = "����<>" & lng��ǰ����
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        If InStr("," & rsTmp!����IDs & ",", "," & lng����ID & ",") > 0 Then
            If blnYes = False Then
                If MsgBox("""" & str�������� & """�������������Ѿ����ڣ��Ƿ�Ҫ��" & """" & str�������� & """�޸�Ϊ��������", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                    blnYes = True
                Else
                    rsTmp.Filter = strFilter
                    Check����ҩ������ = False
                    Exit Function
                End If
            End If
            If blnYes Then
                rsTmp!����IDs = Replace("," & rsTmp!����IDs & ",", "," & lng����ID & ",", ",")
                rsTmp!����IDs = Mid(rsTmp!����IDs, 2)
                If rsTmp!����IDs <> "" Then rsTmp!����IDs = Mid(rsTmp!����IDs, 1, Len(rsTmp!����IDs) - 1)
                rsTmp.Update
            End If
        End If
        rsTmp.MoveNext
    Loop
    rsTmp.Filter = strFilter
    Check����ҩ������ = True
End Function

Private Sub vsUseDept_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim strTmp As String
    
    If KeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call vsUseDept_KeyPress(Index, KeyCode)
    ElseIf KeyCode = vbKeyDelete Then
        With vsUseDept(Index)
            .TextMatrix(.Row, 0) = ""
            .Cell(flexcpData, .Row, 0) = ""
            For i = 0 To .Rows - 1
                If .Cell(flexcpData, i, 0) <> "" Then
                    strTmp = strTmp & "," & .Cell(flexcpData, i, 0)
                End If
            Next
        End With
        If Index = cbo_����ҩ�����Ҷ��շ��� Then
            mrs����ҩ������!����IDs = Mid(strTmp, 2)
            mrs����ҩ������.Update
            Call SetParChange(vsUseDept, Index, mrsPar, True, Get����ҩ������(mrs����ҩ������))
        ElseIf Index = cbo_סԺҩ�����Ҷ��շ��� Then
            mrsסԺҩ������!����IDs = Mid(strTmp, 2)
            mrsסԺҩ������.Update
            Call SetParChange(vsUseDept, Index, mrsPar, True, Get����ҩ������(mrsסԺҩ������))
        ElseIf Index = cbo_סԺ����ִ���Զ���ɷ��� Then
            mrsסԺִ�ж���!����IDs = Mid(strTmp, 2)
            mrsסԺִ�ж���.Update
            Call SetParChange(vsUseDept, Index, mrsPar, True, Get����ҩ������(mrsסԺִ�ж���))
        End If
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0
        If vsUseDept(Index).Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Call EnterNextCell(vsUseDept(Index))
    End If
End Sub

Private Sub vsUseDept_KeyPress(Index As Integer, KeyAscii As Integer)
    vsUseDept(Index).ComboList = "" 'ʹ��ť״̬��������״̬
End Sub

Private Sub vsUseDept_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsUseDept, Index, mrsPar)
End Sub

Private Sub vsUseDept_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim vPoint As POINTAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset, i As Long, strFind As String, strFindSql As String
    Dim strTmp As String, blnCheck As Boolean
    Dim strWhere As String
    
    If Col = 0 Then
        With vsUseDept(Index)
            strFind = .EditText
            If strFind = "" Then
                vsUseDept_KeyDown Index, vbKeyDelete, 0
                Exit Sub
            End If
            If strFind = .TextMatrix(.Row, Col) Then
                Exit Sub
            End If
            If IsNumeric(strFind) Then
                strFindSql = " And (A.���� like [2] OR A.���� Like [2] or A.����=[1]) "
            Else
                strFindSql = " And (A.���� like [2] OR A.���� Like [2] )"
            End If
            
            If Index = cbo_����ҩ�����Ҷ��շ��� Then
                strWhere = " and �������� in ('�ٴ�','���','����','����','����','Ӫ��') And B.������� in (1,3)"
            ElseIf Index = cbo_סԺҩ�����Ҷ��շ��� Then
                strWhere = " and �������� in ('�ٴ�','���','����','����','����','Ӫ��') And B.������� in (2,3)"
            ElseIf Index = cbo_סԺ����ִ���Զ���ɷ��� Then
                strWhere = " And (  b.�������� = '�ٴ�' And ((b.������� In (2, 3)) Or (b.������� = 1 And Exists (Select 1 From ��λ״����¼ C Where b.����id = c.����id)))" & vbNewLine & _
                            "    Or b.�������� = '����' And b.������� In (1, 2, 3))"
            End If
            
            strSQL = "select distinct ID,����,����" & _
                " from ���ű� A,��������˵�� B" & _
                " where A.ID=B.����ID " & strWhere & _
                "       and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & strFindSql & _
                " order by ����"
            
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strFind, "%" & UCase(strFind) & "%")
                    
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    For i = 0 To .Rows - 1
                        If .Cell(flexcpData, i, Col) = rsTmp!ID & "" And i <> Row Then
                            MsgBox "��ǰ�б����Ѿ�����""" & rsTmp!���� & """�ˡ�", vbInformation, Me.Caption
                            Cancel = True
                            Exit Sub
                        End If
                    Next
                    If Index = cbo_����ҩ�����Ҷ��շ��� Then
                        blnCheck = Check����ҩ������(mrs����ҩ������, Val(mrs����ҩ������!���� & ""), Val(rsTmp!ID & ""), rsTmp!���� & "")
                    ElseIf Index = cbo_סԺҩ�����Ҷ��շ��� Then
                        blnCheck = Check����ҩ������(mrsסԺҩ������, Val(mrsסԺҩ������!���� & ""), Val(rsTmp!ID & ""), rsTmp!���� & "")
                    ElseIf Index = cbo_סԺ����ִ���Զ���ɷ��� Then
                        blnCheck = Check����ҩ������(mrsסԺִ�ж���, Val(mrsסԺִ�ж���!���� & ""), Val(rsTmp!ID & ""), rsTmp!���� & "")
                    End If
                    If blnCheck Then
                        .TextMatrix(.Row, Col) = rsTmp!���� & ""
                        .Cell(flexcpData, .Row, Col) = rsTmp!ID & ""
                        .EditText = rsTmp!���� & ""
                        For i = 0 To .Rows - 1
                            If .Cell(flexcpData, i, Col) <> "" Then
                                strTmp = strTmp & "," & .Cell(flexcpData, i, Col)
                            End If
                        Next
                        If Index = cbo_����ҩ�����Ҷ��շ��� Then
                            mrs����ҩ������!����IDs = Mid(strTmp, 2)
                            mrs����ҩ������.Update
                            Call SetParChange(vsUseDept, Index, mrsPar, True, Get����ҩ������(mrs����ҩ������))
                        ElseIf Index = cbo_סԺҩ�����Ҷ��շ��� Then
                            mrsסԺҩ������!����IDs = Mid(strTmp, 2)
                            mrsסԺҩ������.Update
                            Call SetParChange(vsUseDept, Index, mrsPar, True, Get����ҩ������(mrsסԺҩ������))
                        ElseIf Index = cbo_סԺ����ִ���Զ���ɷ��� Then
                            mrsסԺִ�ж���!����IDs = Mid(strTmp, 2)
                            mrsסԺִ�ж���.Update
                            Call SetParChange(vsUseDept, Index, mrsPar, True, Get����ҩ������(mrsסԺִ�ж���))
                        End If
                    Else
                        Cancel = True
                    End If
                Else
                    MsgBox "û���ҵ�ƥ��Ŀ��ҡ�", vbInformation, Me.Caption
                    Cancel = True
                End If
            End If
        End With
    End If
End Sub

Private Function Get����ҩ������(ByRef rsTmp As Recordset) As String
    Dim strValue As String, strFilter As String
    
    If rsTmp.RecordCount = 0 Then Exit Function
    strFilter = rsTmp.Filter
    rsTmp.Filter = 0
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        If rsTmp!����IDs <> "" Then
            strValue = strValue & ";" & rsTmp!����IDs
        End If
        rsTmp.MoveNext
    Loop
    rsTmp.Filter = strFilter
    Get����ҩ������ = Mid(strValue, 2)
End Function

Private Sub setDepartSign()
'���ܣ�����������vsfDepartSign�ؼ�
    Dim strSQL As String, strTmp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    vsfDepartSign.Rows = 1
    
    mrsPar.Filter = "ģ��=" & p�����ڲ����� & " And ������='ǩ��ʹ��ͼƬ'"
    If mrsPar.RecordCount <= 0 Then Exit Sub

    strTmp = zlCommFun.NVL(mrsPar!����ֵ)
    
    On Error GoTo ErrHandle
    strSQL = "Select b.Id As ID, b.���� As ����, b.���� As ����, a.����ֵ As ���� " & _
                  "From Zldeptparas A, ���ű� B " & _
                  "Where a.����id = b.Id And b.�ϼ�id Is Not Null And " & _
                  "a.����id in (Select max(ID) From zlParameters Where ϵͳ = 100 And ģ�� = 1070 And ������ = 'ǩ��ʹ��ͼƬ') " & _
                  "order by ID "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡʹ��ͼƬǩ������", strTmp, zl9ComLib.gstrNodeNo)
    Do While rsTemp.EOF = False
        With vsfDepartSign
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, .ColIndex("ID")) = rsTemp!ID
            .TextMatrix(i, .ColIndex("����")) = rsTemp!���� & ""
            .TextMatrix(i, .ColIndex("����")) = rsTemp!���� & ""
            .TextMatrix(i, .ColIndex("����")) = rsTemp!���� & ""
                        .RowData(i) = NVL(rsTemp!����, 0)
        End With
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    vsfDepartSign.Rows = vsfDepartSign.Rows + 1
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetWaittingMixDept()
'���ܣ�����������vsfWaittingMixDept�ؼ�

    Dim strSQL As String, strTmp As String
    Dim rsTemp As ADODB.Recordset
    Dim l As Long
    
    vsfWaittingMixDept.Rows = 1
    
    mrsPar.Filter = "ģ��=" & p������Һ���� & " And ������='����Һ�����б�'"
    If mrsPar.RecordCount <= 0 Then Exit Sub
    
    strTmp = zlCommFun.NVL(mrsPar!����ֵ)
    
    On Error GoTo ErrHandle
    strSQL = "Select Distinct a.Id, a.����, a.����, Decode(Nvl(c.Column_Value, 0), 0, 0, -1) ���� " & vbNewLine & _
             "From ���ű� A, ��������˵�� B, Table(f_Num2list([1], ',')) C " & vbNewLine & _
             "Where b.����id = a.Id And a.Id = c.Column_Value(+) And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & vbNewLine & _
             "    And b.������� In (1, 3) And b.�������� In ('����', '�ٴ�') And (a.վ�� = [2] Or a.վ�� Is Null) " & vbNewLine & _
             "Order By a.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Һ����", strTmp, zl9ComLib.gstrNodeNo)
    Do While rsTemp.EOF = False
        With vsfWaittingMixDept
            .Rows = .Rows + 1
            l = .Rows - 1
            .TextMatrix(l, .ColIndex("ID")) = rsTemp!ID
            .TextMatrix(l, .ColIndex("����")) = rsTemp!����
            .TextMatrix(l, .ColIndex("����")) = rsTemp!����
            .TextMatrix(l, .ColIndex("����")) = rsTemp!����
        End With
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitNurseItem()
'�����ļ������ؼ�����
    
    '����ͼ
    cbo(cbo_��������˳��).Clear
    cbo(cbo_��������˳��).AddItem "0-����ʾ"
    cbo(cbo_��������˳��).AddItem "1-��ʾ���߼�ͷ"
    cbo(cbo_��������˳��).AddItem "2-��ʾʵ�߼�ͷ"
    cbo(cbo_��������˳��).ListIndex = 0
    
    cbo(cbo_��¶�½�˳��).Clear
    cbo(cbo_��¶�½�˳��).AddItem "0-����ʾ"
    cbo(cbo_��¶�½�˳��).AddItem "1-��ʾ���߼�ͷ"
    cbo(cbo_��¶�½�˳��).AddItem "2-��ʾʵ�߼�ͷ"
    cbo(cbo_��¶�½�˳��).ListIndex = 0
    
    '73309:������,2014-06-24
    cbo(cbo_���������쳣��).Clear
    cbo(cbo_���������쳣��).AddItem "0-����ʾ"
    cbo(cbo_���������쳣��).AddItem "1-��ʾ���߼�ͷ"
    cbo(cbo_���������쳣��).AddItem "2-��ʾʵ�߼�ͷ"
    cbo(cbo_���������쳣��).ListIndex = 0
    
    cbo(cbo_��¶�½��쳣��).Clear
    cbo(cbo_��¶�½��쳣��).AddItem "0-����ʾ"
    cbo(cbo_��¶�½��쳣��).AddItem "1-��ʾ���߼�ͷ"
    cbo(cbo_��¶�½��쳣��).AddItem "2-��ʾʵ�߼�ͷ"
    cbo(cbo_��¶�½��쳣��).AddItem "3-��ʾֱ������"
    cbo(cbo_��¶�½��쳣��).ListIndex = 0
    
    cbo(cbo_������־����).Clear
    cbo(cbo_������־����).AddItem "0-����ʾ"
    cbo(cbo_������־����).AddItem "1-��ʾ����"
    cbo(cbo_������־����).AddItem "2-��ʾ��������"
    cbo(cbo_������־����).ListIndex = 0
    
    cbo(cbo_������־λ��).Clear
    cbo(cbo_������־λ��).AddItem "0-��������"
    cbo(cbo_������־λ��).AddItem "1-��¶�½�"
    cbo(cbo_������־λ��).ListIndex = 0
    
    cbo(cbo_��������ʾ).Clear
    cbo(cbo_��������ʾ).AddItem "0-����"
    cbo(cbo_��������ʾ).AddItem "1-ʵ��"
    cbo(cbo_��������ʾ).ListIndex = 0
    
    cbo(cbo_�쳣����ʾ).Clear
    cbo(cbo_�쳣����ʾ).AddItem "0-����"
    cbo(cbo_�쳣����ʾ).AddItem "1-ʵ��"
    cbo(cbo_�쳣����ʾ).ListIndex = 0
    
    '73309:������,2014-06-24
    cbo(cbo_������״ε�����).Clear
    cbo(cbo_������״ε�����).AddItem "0-������"
    cbo(cbo_������״ε�����).AddItem "1-����������"
    cbo(cbo_������״ε�����).AddItem "2-��ʵ������"
    cbo(cbo_������״ε�����).ListIndex = 0
    
    '��¼��
    '43588,������,2012-09-13,��Ӽ�¼����ǩģʽ
    cbo(cbo_��ǩģʽ).Clear
    cbo(cbo_��ǩģʽ).AddItem "0-Ƹ��ְ��+��ǩȨ��"
    cbo(cbo_��ǩģʽ).AddItem "1-��ǩȨ��"
    cbo(cbo_��ǩģʽ).ListIndex = 0
    
    cbo(cbo_С��ȱʡ��ʶ).Clear
    cbo(cbo_С��ȱʡ��ʶ).AddItem "0-������"
    cbo(cbo_С��ȱʡ��ʶ).AddItem "1-���»����߱�ʶ"
    cbo(cbo_С��ȱʡ��ʶ).AddItem "2-����ֵ�·���˫���߱�ʶ"
    cbo(cbo_С��ȱʡ��ʶ).AddItem "3-�Ϸ������߱�ʶ"
    '72664:������,2014-07-18,���С���ʶ
    cbo(cbo_С��ȱʡ��ʶ).AddItem "4-����ֵ�·��������߱�ʶ"
    cbo(cbo_С��ȱʡ��ʶ).ListIndex = 0
    
    '58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
    cbo(cbo_ǩ������ʾģʽ).Clear
    cbo(cbo_ǩ������ʾģʽ).AddItem "0-��������ʾ"
    cbo(cbo_ǩ������ʾģʽ).AddItem "1-������ʾ"
    cbo(cbo_ǩ������ʾģʽ).AddItem "2-��β����ʾ"
    cbo(cbo_ǩ������ʾģʽ).AddItem "3-β����ʾ"
    cbo(cbo_ǩ������ʾģʽ).ListIndex = 0
    
    '���µ�
    cbo(cbo_��Ժ�Զ���־).Clear
    cbo(cbo_��Ժ�Զ���־).AddItem "0-����ʾ"
    cbo(cbo_��Ժ�Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_��Ժ�Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_��Ժ�Զ���־).ListIndex = 0
    
    cbo(cbo_����Զ���־).Clear
    cbo(cbo_����Զ���־).AddItem "0-����ʾ"
    cbo(cbo_����Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_����Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_����Զ���־).ListIndex = 0
    
    cbo(cbo_ת���Զ���־).Clear
    cbo(cbo_ת���Զ���־).AddItem "0-����ʾ"
    cbo(cbo_ת���Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_ת���Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_ת���Զ���־).AddItem "3-��ʾ˵���Ϳ���"
    cbo(cbo_ת���Զ���־).AddItem "4-��ʾ˵��,����,ʱ��"
    cbo(cbo_ת���Զ���־).ListIndex = 0
    
    cbo(cbo_�����Զ���־).Clear
    cbo(cbo_�����Զ���־).AddItem "0-����ʾ"
    cbo(cbo_�����Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_�����Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_�����Զ���־).ListIndex = 0
    
    cbo(cbo_�����Զ���־).Clear
    cbo(cbo_�����Զ���־).AddItem "0-����ʾ"
    cbo(cbo_�����Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_�����Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_�����Զ���־).ListIndex = 0
    
    cbo(cbo_��Ժ�Զ���־).Clear
    cbo(cbo_��Ժ�Զ���־).AddItem "0-����ʾ"
    cbo(cbo_��Ժ�Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_��Ժ�Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_��Ժ�Զ���־).ListIndex = 0
    
    cbo(cbo_�����Զ���־).Clear
    cbo(cbo_�����Զ���־).AddItem "0-����ʾ"
    cbo(cbo_�����Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_�����Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_�����Զ���־).ListIndex = 0
    
    cbo(cbo_�����Զ���־).Clear
    cbo(cbo_�����Զ���־).AddItem "0-����ʾ"
    cbo(cbo_�����Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_�����Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_�����Զ���־).ListIndex = 0
    
    cbo(cbo_�����Զ���־).Clear
    cbo(cbo_�����Զ���־).AddItem "0-����ʾ"
    cbo(cbo_�����Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_�����Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_�����Զ���־).ListIndex = 0
    
    '73235
    cbo(cbo_ת�����Զ���־).Clear
    cbo(cbo_ת�����Զ���־).AddItem "0-����ʾ"
    cbo(cbo_ת�����Զ���־).AddItem "1-��ʾ˵��"
    cbo(cbo_ת�����Զ���־).AddItem "2-��ʾ˵����ʱ��"
    cbo(cbo_ת�����Զ���־).AddItem "3-��ʾ˵���Ͳ���"
    cbo(cbo_ת�����Զ���־).AddItem "4-��ʾ˵��,����,ʱ��"
    cbo(cbo_ת�����Զ���־).ListIndex = 0
    
    cbo(cbo_��־˵����ʱ�����ӷ���).Clear
    cbo(cbo_��־˵����ʱ�����ӷ���).AddItem "����"
    cbo(cbo_��־˵����ʱ�����ӷ���).AddItem "��"
    cbo(cbo_��־˵����ʱ�����ӷ���).AddItem "�ո�"
    cbo(cbo_��־˵����ʱ�����ӷ���).AddItem ""
    cbo(cbo_��־˵����ʱ�����ӷ���).ListIndex = 0
    
    '51338,������,2012-07-06
    cbo(cbo_��������ȱʡ��ʽ).Clear
    cbo(cbo_��������ȱʡ��ʽ).AddItem "0-����ʾ"
    cbo(cbo_��������ȱʡ��ʽ).AddItem "1-��ʾ0"
    cbo(cbo_��������ȱʡ��ʽ).AddItem "2-��ʾ��������"
    cbo(cbo_��������ȱʡ��ʽ).AddItem "3-��ʾ���ָ�ʽ"
    cbo(cbo_��������ȱʡ��ʽ).ListIndex = 0
    
    '51512,������,2012-07-11
    cbo(cbo_δ��˵����ʾλ��).Clear
    cbo(cbo_δ��˵����ʾλ��).AddItem "0-��ʾ������"
    cbo(cbo_δ��˵����ʾλ��).AddItem "1-��ʾ������"
    cbo(cbo_δ��˵����ʾλ��).AddItem "2-����ʾ"
    cbo(cbo_δ��˵����ʾλ��).ListIndex = 0
    
    cbo(cbo_���²�����ʾ��ʽ).Clear
    cbo(cbo_���²�����ʾ��ʽ).AddItem "0-��ͷ"
    cbo(cbo_���²�����ʾ��ʽ).AddItem "1-����"
    cbo(cbo_���²�����ʾ��ʽ).AddItem "2-����+��ͷ"
    cbo(cbo_���²�����ʾ��ʽ).AddItem "3-����+����"
    cbo(cbo_���²�����ʾ��ʽ).ListIndex = 0
    
    '73316:������,2014-06-26
    cbo(cbo_������������ʾλ��).Clear
    cbo(cbo_������������ʾλ��).AddItem "0-�������д����������R(ȱʡ��ʽ)"
    cbo(cbo_������������ʾλ��).AddItem "1-�������д����Ƶ��,��Ӧʱ������������Ϸ������������������,�á���ʶ��ʼ����ʶ��ֹ"
    cbo(cbo_������������ʾλ��).AddItem "2-�������дA+����ֵ"
    zlControl.CboSetWidth cbo(cbo_������������ʾλ��).hwnd, 7500
    '72663:������,2014-08-08
    cbo(cbo_�������������ʾλ��).Clear
    cbo(cbo_�������������ʾλ��).AddItem "0-���ϵ���(�����ݼ̳�)"
    cbo(cbo_�������������ʾλ��).AddItem "1-���ϵ���(�����ݼ̳�)"
    cbo(cbo_�������������ʾλ��).AddItem "2-���µ���(�����ݼ̳�)"
    cbo(cbo_�������������ʾλ��).AddItem "3-���µ���(�����ݼ̳�)"
End Sub

Private Sub SetCOLOR(vData As OLE_COLOR, ByVal Index As Integer)
    Dim lRow As Long, lCol As Long
    shpValue(Index).Visible = True
    Select Case CStr(Hex(vData))
    Case "0"
        lblColor(Index) = "��ɫ"
        lRow = 0
        lCol = 0
    Case "3399"
        lblColor(Index) = "��ɫ"
        lRow = 0
        lCol = 1
    Case "3333"
        lblColor(Index) = "���ɫ"
        lRow = 0
        lCol = 2
    Case "3300"
        lblColor(Index) = "����"
        lRow = 0
        lCol = 3
    Case "663300"
        lblColor(Index) = "����"
        lRow = 0
        lCol = 4
    Case "800000"
        lblColor(Index) = "����"
        lRow = 0
        lCol = 5
    Case "993333"
        lblColor(Index) = "����"
        lRow = 0
        lCol = 6
    Case "333333"
        lblColor(Index) = "��ɫ-80%"
        lRow = 0
        lCol = 7
    Case "80"
        lblColor(Index) = "���"
        lRow = 1
        lCol = 0
    Case "66FF"
        lblColor(Index) = "��ɫ"
        lRow = 1
        lCol = 1
    Case "8080"
        lblColor(Index) = "���"
        lRow = 1
        lCol = 2
    Case "8000"
        lblColor(Index) = "��ɫ"
        lRow = 1
        lCol = 3
    Case "808000"
        lblColor(Index) = "��ɫ"
        lRow = 1
        lCol = 4
    Case "FF0000"
        lblColor(Index) = "��ɫ"
        lRow = 1
        lCol = 5
    Case "996666"
        lblColor(Index) = "��-��"
        lRow = 1
        lCol = 6
    Case "808080"
        lblColor(Index) = "��ɫ-50%"
        lRow = 1
        lCol = 7
    Case "FF"
        lblColor(Index) = "��ɫ"
        lRow = 2
        lCol = 0
    Case "99FF"
        lblColor(Index) = "ǳ��ɫ"
        lRow = 2
        lCol = 1
    Case "CC99"
        lblColor(Index) = "���ɫ"
        lRow = 2
        lCol = 2
    Case "669933"
        lblColor(Index) = "����"
        lRow = 2
        lCol = 3
    Case "CCCC33"
        lblColor(Index) = "ˮ��ɫ"
        lRow = 2
        lCol = 4
    Case "FF6633"
        lblColor(Index) = "ǳ��"
        lRow = 2
        lCol = 5
    Case "800080"
        lblColor(Index) = "������"
        lRow = 2
        lCol = 6
    Case "999999"
        lblColor(Index) = "��ɫ-40%"
        lRow = 2
        lCol = 7
    Case "FF00FF"
        lblColor(Index) = "�ۺ�"
        lRow = 3
        lCol = 0
    Case "CCFF"
        lblColor(Index) = "��ɫ"
        lRow = 3
        lCol = 1
    Case "FFFF"
        lblColor(Index) = "��ɫ"
        lRow = 3
        lCol = 2
    Case "FF00"
        lblColor(Index) = "����"
        lRow = 3
        lCol = 3
    Case "FFFF00"
        lblColor(Index) = "����"
        lRow = 3
        lCol = 4
    Case "FFCC00"
        lblColor(Index) = "����"
        lRow = 3
        lCol = 5
    Case "663399"
        lblColor(Index) = "÷��"
        lRow = 3
        lCol = 6
    Case "C0C0C0"
        lblColor(Index) = "��ɫ-25%"
        lRow = 3
        lCol = 7
    Case "CC99FF"
        lblColor(Index) = "õ���"
        lRow = 4
        lCol = 0
    Case "99CCFF"
        lblColor(Index) = "��ɫ"
        lRow = 4
        lCol = 1
    Case "99FFFF"
        lblColor(Index) = "ǳ��"
        lRow = 4
        lCol = 2
    Case "CCFFCC"
        lblColor(Index) = "ǳ��"
        lRow = 4
        lCol = 3
    Case "FFFFCC"
        lblColor(Index) = "ǳ����"
        lRow = 4
        lCol = 4
    Case "FFCC99"
        lblColor(Index) = "����"
        lRow = 4
        lCol = 5
    Case "FF99CC"
        lblColor(Index) = "����"
        lRow = 4
        lCol = 6
    Case "FFFFFF"
        lblColor(Index) = "��ɫ"
        lRow = 4
        lCol = 7
    Case Else
        lblColor(Index) = "&H" & CStr(Hex(picColor(Index).BackColor))
    End Select
    shpBorder(Index).Visible = False
    shpValue(Index).Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    shpValue(Index).Visible = True
    If vData = -9999997 Or vData = -1 Then
    
    Else
        picColor(Index).BackColor = vData
    End If
End Sub

Private Sub Set���뵥���û���(ByVal strPar As String)
'���ܣ����ý���ؼ� �����������뵥���û���
    Dim strTmp As String
    Dim varTmp As Variant
    On Error GoTo errH
    varTmp = Split(strPar, "|")
    strTmp = varTmp(0)
    chk(chk_���뵥���û���������).value = Mid(strTmp, 1, 1)
    chk(chk_���뵥���û���סԺ���).value = Mid(strTmp, 2, 1)
    strTmp = varTmp(1)
    chk(chk_���뵥���û����������).value = Mid(strTmp, 1, 1)
    chk(chk_���뵥���û���סԺ����).value = Mid(strTmp, 2, 1)
    strTmp = varTmp(2)
    chk(chk_���뵥���û���������Ѫ).value = Mid(strTmp, 1, 1)
    chk(chk_���뵥���û���סԺ��Ѫ).value = Mid(strTmp, 2, 1)
    strTmp = varTmp(3)
    chk(chk_���뵥���û�����������).value = Mid(strTmp, 1, 1)
    chk(chk_���뵥���û���סԺ����).value = Mid(strTmp, 2, 1)
    strTmp = varTmp(4)
    chk(chk_���뵥���û��ڻ���).value = Mid(strTmp, 1, 1)
    Exit Sub
errH:
    MsgBox "ϵͳ���������뵥���û��ڣ�����ֵ��ʽ���ԣ�", vbInformation, gstrSysName
    Err.Clear
End Sub

Private Function Get���뵥���û���() As String
'���ܣ��ӽ����ȡ����ֵ
    Dim strTmp As String
    strTmp = chk(chk_���뵥���û���������).value & chk(chk_���뵥���û���סԺ���).value & "|" & _
    chk(chk_���뵥���û����������).value & chk(chk_���뵥���û���סԺ����).value & "|" & _
    chk(chk_���뵥���û���������Ѫ).value & chk(chk_���뵥���û���סԺ��Ѫ).value & "|" & _
    chk(chk_���뵥���û�����������).value & chk(chk_���뵥���û���סԺ����).value & "|" & _
    chk(chk_���뵥���û��ڻ���).value
    Get���뵥���û��� = strTmp
End Function

Private Sub InitRsҩ������(ByRef rsIn As ADODB.Recordset)
'���ܣ���ʼ������ҩ����Ӧ��ϵ��¼��
    Set rsIn = New ADODB.Recordset
    rsIn.Fields.Append "����", adVarChar, 1000
    rsIn.Fields.Append "����IDs", adVarChar, 400000
    rsIn.Fields.Append "���÷��ϲ���", adVarChar, 40000
    rsIn.Fields.Append "������ҩ��", adVarChar, 40000
    rsIn.Fields.Append "���ó�ҩ��", adVarChar, 40000
    rsIn.Fields.Append "������ҩ��", adVarChar, 40000
    rsIn.Fields.Append "ȱʡ���ϲ���", adVarChar, 40000
    rsIn.Fields.Append "ȱʡ��ҩ��", adVarChar, 40000
    rsIn.Fields.Append "ȱʡ��ҩ��", adVarChar, 40000
    rsIn.Fields.Append "ȱʡ��ҩ��", adVarChar, 40000
'    rsIn.Fields.Append "���ϲ��Ŵ���", adVarChar, 40000
'    rsIn.Fields.Append "��ҩ������", adVarChar, 40000
'    rsIn.Fields.Append "��ҩ������", adVarChar, 40000
'    rsIn.Fields.Append "��ҩ������", adVarChar, 40000
    rsIn.CursorLocation = adUseClient
    rsIn.LockType = adLockOptimistic
    rsIn.CursorType = adOpenStatic
    rsIn.Open
End Sub

Private Function isPiotBlood() As Boolean
'�Ƿ�����Ѫ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHand
    strSQL = "Select 1 From zlSystems Where ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 2200)
    isPiotBlood = (Not rsTmp.EOF)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SaveסԺִ�ж���()
'���ܣ����� סԺ����ִ���Զ���ɷ���
    Dim i As Long, strFilter As String
    Dim str���� As String
    
    If Frame14.Tag <> "���޸�" Then
        mrsPar.Filter = "������='סԺ����ִ���Զ���ɷ���' and �޸�״̬=1 "
        If Not mrsPar.EOF Then Frame14.Tag = "���޸�"
    End If
    
    If Frame14.Tag = "���޸�" Then
        str���� = ""
        Call zlDatabase.DelDeptPara("����ִ���Զ����ҽ�����", glngSys, pסԺҽ������)
        strFilter = mrsסԺִ�ж���.Filter
        mrsסԺִ�ж���.Filter = 0
        If mrsסԺִ�ж���.RecordCount > 0 Then mrsסԺִ�ж���.MoveFirst
        Do While Not mrsסԺִ�ж���.EOF
            If mrsסԺִ�ж���!����IDs & "" <> "" Then
                str���� = str���� & ";" & mrsסԺִ�ж���!����IDs
                For i = 0 To UBound(Split(mrsסԺִ�ж���!����IDs, ","))
                    Call zlDatabase.SetPara("����ִ���Զ����ҽ�����", mrsסԺִ�ж���!ҽ����� & "", glngSys, pסԺҽ������, , Split(mrsסԺִ�ж���!����IDs, ",")(i))
                Next
            End If
            mrsסԺִ�ж���.MoveNext
        Loop
        Call zlDatabase.SetPara("סԺ����ִ���Զ���ɷ���", Mid(str����, 2), glngSys, pסԺҽ������)
        mrsסԺִ�ж���.Filter = IIF(strFilter = "0", 0, strFilter)
        Frame14.Tag = ""
    End If
End Sub

Private Sub SetValueסԺִ��()
'���ܣ�סԺ����ִ���Զ���ɷ��� ĳ��������ֵ
    Dim strTmp As String
    Dim i As Long
    For i = 0 To lst(lst_����ִ���Զ����ҽ�����).ListCount - 1
        If lst(lst_����ִ���Զ����ҽ�����).Selected(i) Then
            strTmp = strTmp & "," & i
        End If
    Next
    mrsסԺִ�ж���!ҽ����� = Mid(strTmp, 2)
End Sub

Private Sub Set�Զ���ɷ���(ByVal lngIndex As Long)
    Dim str��� As String
    Dim i As Long
    Dim strDeptIDs As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    mrsסԺִ�ж���.Filter = "����=" & cbo(lngIndex).ItemData(cbo(lngIndex).ListIndex)
    If mrsסԺִ�ж���.RecordCount > 0 Then
        str��� = mrsסԺִ�ж���!ҽ����� & ""
        strDeptIDs = mrsסԺִ�ж���!����IDs & ""
    End If
    
    
    strSQL = "select ID,���� From ���ű� Where ID in(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDeptIDs)
    vsUseDept(lngIndex).Enabled = True
    vsUseDept(lngIndex).Rows = 0
    Do While Not rsTmp.EOF
        vsUseDept(lngIndex).Rows = vsUseDept(lngIndex).Rows + 1
        vsUseDept(lngIndex).TextMatrix(vsUseDept(lngIndex).Rows - 1, 0) = rsTmp!���� & ""
        vsUseDept(lngIndex).Cell(flexcpData, vsUseDept(lngIndex).Rows - 1, 0) = rsTmp!ID & ""
        rsTmp.MoveNext
    Loop
    
    If vsUseDept(lngIndex).Rows = 0 Then
        vsUseDept(lngIndex).Rows = 1
    End If
    
    vsUseDept(lngIndex).Enabled = True
    lst(lst_����ִ���Զ����ҽ�����).Enabled = True
    If str��� <> "" Then
        If str��� = "*" Then str��� = "012345678"
        For i = 0 To 8
            lst(lst_����ִ���Զ����ҽ�����).Selected(i) = InStr(str���, i) > 0
        Next
    End If
End Sub
