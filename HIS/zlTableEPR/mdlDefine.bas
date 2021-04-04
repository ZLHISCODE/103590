Attribute VB_Name = "mdlDefine"
Option Explicit

'################################################################################################################
'##     �˵�
'################################################################################################################
'���˵�
Public Const ID_File_Menu = 101 '�ļ�
Public Const ID_Edit_Menu = 102 '�༭
Public Const ID_Insert_Menu = 103 '����
Public Const ID_Com_Bar = 104   '����
Public Const ID_Sign_Bar = 105  'ǩ��
Public Const ID_Format_Bar = 106 '��ʽ
Public Const ID_Table_Bar = 107 '�������

'�ļ� "File"
Public Const ID_FILE_CLEAR = 300                '���
Public Const ID_FILE_IMPORT = 301               '����
Public Const ID_FILE_CLOSE = 302                '�ر� ��
Public Const ID_FILE_SAVE = 303                 '����
Public Const ID_FILE_SAVEAS = 304               '���Ϊ
Public Const ID_FILE_PAGESETUP = 305            'ҳ������
Public Const ID_FILE_PRINTPREVIEW = 306         '��ӡԤ��
Public Const ID_FILE_PRINT = 307                '��ӡ
Public Const ID_FILE_EXIT = 308                 '�˳�
Public Const ID_FILE_SAVEASEPRDEMO = 309        '���Ϊ����
Public Const ID_FILE_EXPORTTOXML = 310          '����ΪXML�ļ�
Public Const ID_FILE_IMPORTFROMXML = 311        '��XML�ļ�����
Public Const ID_FILE_EXPORTTOHTML = 312         '����ΪHTML�ļ�
Public Const ID_FILE_PRINTINWORD = 313          '��Word�д�ӡ
Public Const ID_FILE_SAVEASSEGMENT = 314        '���ΪƬ��
Public Const ID_FILE_SAVE_QUIT = 315            '���沢�˳�

'�༭ "Edit"
Public Const ID_EDIT_UNDO = 320                 '����
Public Const ID_EDIT_REDO = 321                 '����
Public Const ID_EDIT_CUT = 322                  '����
Public Const ID_EDIT_COPY = 323                 '����
Public Const ID_EDIT_PASTE = 324                'ճ��
Public Const ID_EDIT_DELETE = 325               'ɾ��
Public Const ID_EDIT_SELECTALL = 326            'ȫѡ
Public Const ID_EDIT_FIND = 327                 '����
Public Const ID_EDIT_REPLACE = 328              '�滻
Public Const ID_EDIT_FINDNEXT = 329             '������һ��
Public Const ID_EDIT_FORMATBRUSH = 330          '��ʽˢ
Public Const ID_EDIT_ADDCOMPEND = 331           '�������
Public Const ID_EDIT_MODCOMPEND = 332           '�޸����
Public Const ID_EDIT_DELCOMPEND = 333           'ɾ�����
Public Const ID_EDIT_REFCOMPEND = 334           'ˢ�����
Public Const ID_EDIT_SAVEASPHRASE = 335         '��Ϊ�ʾ�ʾ��
Public Const ID_EDIT_COMPENDWORD = 336          '��ٴʾ����

Public Const ID_EDIT_MARKEDPIC = 337            '����޸�
Public Const ID_EDIT_OUTERPIC = 338             '��ͼ����
Public Const ID_EDIT_DELETEELEMENT = 339        'ɾ��Ҫ��

'��ͼ "View"
Public Const ID_VIEW_STRUCTURE = 340            '�ĵ��ṹͼ
Public Const ID_VIEW_PHRASEDEMO = 341           '�ʾ�ʾ���б�
Public Const ID_VIEW_SEGMENT = 342              'ʾ��Ƭ���б�
Public Const ID_VIEW_HEADFOOT = 343             'ҳüҳ��
Public Const ID_VIEW_GRID = 344                 '������
Public Const ID_VIEW_PACSPIC = 345              'PACSͼƬ���б���
Public Const ID_VIEW_MULTIDOCVIEW = 346         '���ĵ�����
Public Const ID_VIEW_CHARCOUNT = 347            '����ͳ��
Public Const ID_VIEW_RULER = 348                '���
Public Const ID_VIEW_PENWINDOW = 349            '��д���봰��
Public Const ID_VIEW_HISTORYWINDOW = 3400       '����ҳ������

'���� "Insert"
Public Const ID_INSERT_DATETIME = 350           '����ʱ��
Public Const ID_INSERT_SPECIALCHAR = 351        '�������
Public Const ID_INSERT_PICTURE = 352            'ͼƬ
Public Const ID_INSERT_TABLE = 353              '���
Public Const ID_INSERT_ELEMENT = 354            '����Ҫ��
Public Const ID_INSERT_EPRDEMO = 355            'ȫ��ʾ��
Public Const ID_INSERT_DATE = 356               '��������
Public Const ID_INSERT_TIME = 357               '����ʱ��
Public Const ID_INSERT_DOCADVISE = 358          '���뱾�ξ���ҽ��
Public Const ID_INSERT_AUTORECOGNISE = 359      '����ʶ������Ҫ�ء��ֵ���Ŀ��
Public Const ID_INSERT_PRECOMPEND = 360         '����Ԥ�����
Public Const ID_INSERT_PACSPIC = 361            '����PACSͼƬ��

'��ʽ "Format"
Public Const ID_FORMAT_FONT = 390               '����
Public Const ID_FORMAT_BACKGROUND = 391         '����ɫ
Public Const ID_FORMAT_PROTECT = 392            '����
Public Const ID_FORMAT_BOLD = 393               '����
Public Const ID_FORMAT_ITALIC = 394             'б��
Public Const ID_FORMAT_SUPER = 395              '�ϱ�
Public Const ID_FORMAT_SUB = 396                '�±�
Public Const ID_FORMAT_UNDERLINE_THIN = 397     '�»��ߣ�ϸ��
Public Const ID_FORMAT_UNDERLINE_THICK = 398    '�»��ߣ�����
Public Const ID_FORMAT_UNDERLINE_WAVE = 399     '�»��ߣ�������
Public Const ID_FORMAT_UNDERLINE_DOT = 400      '�»��ߣ�����
Public Const ID_FORMAT_UNDERLINE_DASH = 401     '�»��ߣ�����
Public Const ID_FORMAT_UNDERLINE_DASHDOT = 402  '�»��ߣ��㻮��
Public Const ID_FORMAT_UNDERLINE_DASHDOT2 = 403 '�»��ߣ�˫�㻮��
Public Const ID_FORMAT_ALIGNLEFT = 404          '���뷽ʽ�������
Public Const ID_FORMAT_ALIGNCENTER = 405        '���뷽ʽ�������
Public Const ID_FORMAT_ALIGNRIGHT = 406         '���뷽ʽ�������
Public Const ID_FORMAT_LISTNONE = 407           '��Ŀ���ţ���
Public Const ID_FORMAT_LISTBULLETS = 408        '��Ŀ���ţ���Ŀ����
Public Const ID_FORMAT_LISTLCHAR = 409          '��Ŀ���ţ�Сд��ĸ
Public Const ID_FORMAT_LISTUCHAR = 410          '��Ŀ���ţ���д��ĸ
Public Const ID_FORMAT_LISTLROME = 411          '��Ŀ���ţ�Сд��������
Public Const ID_FORMAT_LISTUROME = 412          '��Ŀ���ţ���д��������
Public Const ID_FORMAT_LINESPACE = 413          '�м��
Public Const ID_FORMAT_SPACEBEFORE = 414        '��ǰ����
Public Const ID_FORMAT_SPACEAFTER = 415         '�κ����
Public Const ID_FORMAT_FIRSTINDENT = 416        '��������
Public Const ID_FORMAT_FIRSTHUNGING = 417       '��������
Public Const ID_FORMAT_INDENTDECREASE = 418     '����������
Public Const ID_FORMAT_INDENTINCREASE = 419     '����������
Public Const ID_FORMAT_UNDERLINE = 420          '�»���
Public Const ID_FORMAT_LISTARABIC = 421         '��Ŀ���ţ�����������
Public Const ID_FORMAT_PARA = 422               '��������
Public Const ID_FORMAT_LINESPACE1 = 423         '�м�ࣺ1.0��
Public Const ID_FORMAT_LINESPACE2 = 424         '�м�ࣺ1.3��
Public Const ID_FORMAT_LINESPACE3 = 425         '�м�ࣺ1.5��
Public Const ID_FORMAT_LINESPACE4 = 426         '�м�ࣺ2.0��
Public Const ID_FORMAT_LINESPACE5 = 427         '�м�ࣺ2.5��
Public Const ID_FORMAT_LINESPACE6 = 428         '�м�ࣺ3.0��
Public Const ID_FORMAT_LINESPACE7 = 429         '�м�ࣺ����...
Public Const ID_FORMAT_HIGHLIGHT = 530          '������ʾ ��
Public Const ID_FORMAT_FORECOLOR = 531          '������ɫ
Public Const ID_FORMAT_STYLE = 532              '������ʽ
Public Const ID_FORMAT_FONTNAME = 533           '��������
Public Const ID_FORMAT_FONTSIZE = 534           '����ߴ�
Public Const ID_FORMAT_UNDERLINE_NONE = 535     '�»��ߣ���
Public Const ID_FORMAT_LISTSETUP = 536          '��Ŀ��������
Public Const ID_FORMAT_STYLEWINDOW = 537        '��ʽ����

'��� "Table"
Public Const ID_TABLE_INSERTTABLE = 430         '������ ��
Public Const ID_TABLE_INSERTCOLLEFT = 431       '�����У���ߣ�
Public Const ID_TABLE_INSERTCOLRIGHT = 432      '�����У��ұߣ�
Public Const ID_TABLE_INSERTROWUP = 433         '�����У����ϣ�
Public Const ID_TABLE_INSERTROWDOWN = 434       '�����У����£�
Public Const ID_TABLE_INSERTCELL = 435          '���뵥Ԫ��...
Public Const ID_TABLE_DELETETABLE = 436         'ɾ����� ��
Public Const ID_TABLE_DELETECOL = 437           'ɾ����
Public Const ID_TABLE_DELETEROW = 438           'ɾ����
Public Const ID_TABLE_DELETECELL = 439          'ɾ����Ԫ��
Public Const ID_TABLE_FORMATCELL = 440          '��Ԫ���ʽ
Public Const ID_TABLE_FORMATROWHEIGHT = 441     '�и�
Public Const ID_TABLE_FORMATCOLWIDTH = 442      '�п�
Public Const ID_TABLE_INSERTPICTURE = 443       '����ͼƬ
Public Const ID_TABLE_BEELEMENTS = 444          '��������Ҫ��
Public Const ID_TABLE_MERGE = 445               '�ϲ���Ԫ��
Public Const ID_TABLE_CELLALIGNMENT1 = 446      '��Ԫ����뷽ʽ
Public Const ID_TABLE_CELLALIGNMENT2 = 447      '��Ԫ����뷽ʽ
Public Const ID_TABLE_CELLALIGNMENT3 = 448      '��Ԫ����뷽ʽ
Public Const ID_TABLE_CELLALIGNMENT4 = 449      '��Ԫ����뷽ʽ
Public Const ID_TABLE_CELLALIGNMENT5 = 450      '��Ԫ����뷽ʽ
Public Const ID_TABLE_CELLALIGNMENT6 = 451      '��Ԫ����뷽ʽ
Public Const ID_TABLE_CELLALIGNMENT7 = 452      '��Ԫ����뷽ʽ
Public Const ID_TABLE_CELLALIGNMENT8 = 453      '��Ԫ����뷽ʽ
Public Const ID_TABLE_CELLALIGNMENT9 = 454      '��Ԫ����뷽ʽ
Public Const ID_TABLE_BORDERSTYLE1 = 455        '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE2 = 456        '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE3 = 457        '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE4 = 458        '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE5 = 459        '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE6 = 460        '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE7 = 461        '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE8 = 462        '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE9 = 463        '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE10 = 464       '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE11 = 465       '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE12 = 466       '��Ԫ��߿���ʽ
Public Const ID_TABLE_BORDERSTYLE13 = 467            '��Ԫ��߿���ʽ
Public Const ID_TABLE_INSERTINHERITROW = 468    '����̳���
Public Const ID_TABLE_INSERTINHERITCOL = 469    '����̳���
Public Const ID_TABLE_CELLPROTECTED = 470       '������Ԫ��

'���� "Help"
Public Const ID_HELP_CONTENT = 500              '��������
Public Const ID_HELP_ASSISTANT = 501            'ϵͳ���� ��
Public Const ID_HELP_CONTACT = 502              '���ͷ���
Public Const ID_HELP_ONLINE = 503               '����ҽҵ
Public Const ID_HELP_ABOUT = 504                '����...
Public Const ID_HELP_WEBFORUM = 505             '������̳(&F)

'#########################################################################
'##     ������������ID
'#########################################################################

'��ͼ������ "Draw"
Public Const ID_DRAW_SELECT = 550               'ѡ��
Public Const ID_DRAW_MOVE = 551                 '�ƶ�
Public Const ID_DRAW_LINE = 552                 'ֱ��
Public Const ID_DRAW_MLINE = 553                '����
Public Const ID_DRAW_RECT = 554                 '����
Public Const ID_DRAW_MRECT = 555                '�����
Public Const ID_DRAW_CIRCLE = 556               '��Բ
Public Const ID_DRAW_TEXT = 557                 '�ı�
Public Const ID_DRAW_DELETE = 558               'ɾ��
Public Const ID_DRAW_UNDO = 559                 'ȡ��
Public Const ID_DRAW_REDO = 560                 '����
Public Const ID_DRAW_RESET = 561                '���
Public Const ID_DRAW_FILLCOLOR = 562            '���ɫ
Public Const ID_DRAW_LINECOLOR = 563            '����ɫ
Public Const ID_DRAW_FONTCOLOR = 564            '����ɫ
Public Const ID_DRAW_FILLSTYLE = 565            '�����ʽ
Public Const ID_DRAW_LINESTYLE = 566            '������ʽ
Public Const ID_DRAW_LINEWIDTH = 567            '�������
Public Const ID_DRAW_FILLNONE = 630             '��䷽ʽ
Public Const ID_DRAW_FILLALL = 631
Public Const ID_DRAW_FILLH = 632
Public Const ID_DRAW_FILLV = 633
Public Const ID_DRAW_FILLHV = 634
Public Const ID_DRAW_FILLR = 635
Public Const ID_DRAW_FILLL = 636
Public Const ID_DRAW_FILLLR = 637
Public Const ID_DRAW_LINECONTINUE = 639         '������ʽ
Public Const ID_DRAW_LINEDOT = 640
Public Const ID_DRAW_LINEDASH = 641
Public Const ID_DRAW_LINEDASHDOT = 642
Public Const ID_DRAW_LINEDASHDOT2 = 643
Public Const ID_DRAW_LINEWIDTH1 = 644           '�������
Public Const ID_DRAW_LINEWIDTH2 = 645
Public Const ID_DRAW_LINEWIDTH3 = 646
Public Const ID_DRAW_LINEWIDTH4 = 647
Public Const ID_DRAW_LINEWIDTH5 = 648
Public Const ID_DRAW_SEQUENCENUMBER = 650       '˳����
Public Const ID_DRAW_CLEARNUMBERS = 651         '���˳����

'��񹤾���
Public Const ID_TABLE_MERGEANDCENTER = 580      '�ϲ�������
Public Const ID_TABLE_SAMECOLWIDTH = 581        '��ͬ�п�
Public Const ID_TABLE_SAMEROWHEIGHT = 582       '��ͬ�и�
Public Const ID_TABLE_CURRENCY = 583            '����
Public Const ID_TABLE_PERCENT = 584             '�ٷֱ�
Public Const ID_TABLE_KILOBIT = 585             'ǧ��λ
Public Const ID_TABLE_DIGITSINCREASE = 586      '����С����
Public Const ID_TABLE_DIGITSDECREASE = 587      '����С����
Public Const ID_TABLE_BORDERSTYLE = 588         '�߿���ʽ
Public Const ID_TABLE_CELLALIGNMENT = 589       '��Ԫ����뷽ʽ
Public Const ID_TABLE_FORMULA = 590             '��ʽ��
Public Const ID_TABLE_INSERTTABLE_BAR = 591     '������
Public Const ID_TABLE_PROPERTY = 592            '�������


'#########################################################################
'##     ���������Ĳ˵�
'#########################################################################
Public Const ID_SIGN = 710                      'ǩ��
Public Const ID_UNTREAD = 711                   '����
Public Const ID_SIGN_QUIT = 712                 'ǩ�����˳��༭����
Public Const ID_REVISION_PREV = 715             'ǰһ���޶�
Public Const ID_REVISION_NEXT = 716             '��һ���޶�
Public Const ID_REVISION_RESET = 717            'ȡ����ѡ�޶�
Public Const ID_DIAGNOSIS = 720                 '���
Public Const ID_ELEMENT_TOSTRING = 722          'ת��Ϊ���ı�
Public Const ID_EDIT_BACKSPACE = 723            '�༭���а�BackSpace��
Public Const ID_ELEMENT_CLEAR = 724             '����ı�
Public Const ID_ELEMENT_UPDATE = 725            '�����ı�
Public Const ID_DesignTest = 9999               '��ƻ����²��԰�Ť

'PACS����
Public Const ID_PACS_DeleteMarkedPic = 900      'ɾ�����ͼ
Public Const ID_PACS_DeletePacsImg = 901        'ɾ��PACS����ͼƬ
Public Const ID_PACS_Layout = 903               '���ֵ���
Public Const ID_PACS_Left = 904                 '���ͼ�����
Public Const ID_PACS_Right = 905                '���ͼ���ұ�
Public Const ID_PACS_None = 906                 '�ޱ��ͼ
