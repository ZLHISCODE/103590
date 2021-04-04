Attribute VB_Name = "mDefinitions"
Option Explicit

'����ע��ϵͳ�ȼ�
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

'ϵͳ���������
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B

'�ļ�
Public Const ID_FILE_OPEN = 300                 '��
Public Const ID_FILE_SAVE = 301                 '����
Public Const ID_FILE_SAVEAS = 302               '���Ϊ
Public Const ID_FILE_PRINT = 303                '��ӡ
Public Const ID_FILE_EXIT = 304                 '�˳�

'�༭
Public Const ID_EDIT_UNDO = 400                 '����
Public Const ID_EDIT_REDO = 401                 '����
Public Const ID_EDIT_COPY = 402                 '����
Public Const ID_EDIT_PASTE = 403                'ճ��
Public Const ID_EDIT_SIZE = 404                 '�����ߴ�
Public Const ID_EDIT_ORIENT = 405               '��������
Public Const ID_EDIT_SCROLLMODE = 406           '��ģʽ
Public Const ID_EDIT_CROPMODE = 407             '����ģʽ

'����
Public Const ID_ZOOM_IN = 500                   '�Ŵ�
Public Const ID_ZOOM_OUT = 501                  '��С
Public Const ID_ZOOM_11 = 502                   '1:1
Public Const ID_ZOOM_FIT = 503                  '�ʺ�

'��ɫ
Public Const ID_COLOR_BLACKWHITE = 600          '�Ҷ�-�ڰ�
Public Const ID_COLOR_GREYS16 = 601             '�Ҷ�-16ɫ
Public Const ID_COLOR_GREYS256 = 602            '�Ҷ�-256ɫ
Public Const ID_COLOR_COLOR2 = 603              '��ɫ-2ɫ
Public Const ID_COLOR_COLOR16 = 604             '��ɫ-16ɫ
Public Const ID_COLOR_COLOR256 = 605            '��ɫ-256ɫ
Public Const ID_COLOR_TRUECOLOR = 606           '���ɫ

'����
Public Const ID_ADJUST_BRIGHT = 700             '����
Public Const ID_ADJUST_CONTRAST = 701           '�Աȶ�
Public Const ID_ADJUST_SITUATION = 702          '���Ͷ�
Public Const ID_ADJUST_FILTERBROWSER = 703      '�˾������

'�˾�
Public Const ID_FILTER_COLOR1 = 800             '��ɫ���Ҷ�
Public Const ID_FILTER_COLOR2 = 801             '��ɫ����ƬЧ��
Public Const ID_FILTER_COLOR3 = 802             '��ɫ������Ƭ
Public Const ID_FILTER_COLOR4 = 803             '��ɫ����ɫ���
Public Const ID_FILTER_COLOR5 = 804             '��ɫ���滻 HS...
Public Const ID_FILTER_COLOR6 = 805             '��ɫ���滻 L...
Public Const ID_FILTER_COLOR7 = 806             '��ɫ���ع����

Public Const ID_FILTER_DEF1 = 810               '�����ȣ�ģ��
Public Const ID_FILTER_DEF2 = 811               '�����ȣ��ữ
Public Const ID_FILTER_DEF3 = 812               '�����ȣ���
Public Const ID_FILTER_DEF4 = 813               '�����ȣ���ɢ
Public Const ID_FILTER_DEF5 = 814               '�����ȣ����ػ�
Public Const ID_FILTER_DEF6 = 815               '�����ȣ�ȥ��
Public Const ID_FILTER_DEF7 = 816               '�����ȣ���һ��ȥ��

Public Const ID_FILTER_EDGES1 = 820             '��Ե������
Public Const ID_FILTER_EDGES2 = 821             '��Ե������
Public Const ID_FILTER_EDGES3 = 822             '��Ե���滭
Public Const ID_FILTER_EDGES4 = 823             '��Ե����Ŀ

Public Const ID_FILTER_SPECIAL1 = 830           '���⣭����
Public Const ID_FILTER_SPECIAL2 = 831           '���⣭ɨ����
Public Const ID_FILTER_SPECIAL3 = 832           '���⣭����
Public Const ID_FILTER_SPECIAL4 = 833           '���⣭��ʴ
Public Const ID_FILTER_SPECIAL5 = 834           '���⣭����...

'��ͼ
Public Const ID_VIEW_TOOLBARLIST = 59392        '�������б�
Public Const ID_VIEW_PANORAMIC = 900            '����ͼ
Public Const ID_VIEW_PROPERTY = 901             '����

'���� "Help"
Public Const ID_HELP_CONTENT = 902              '��������
Public Const ID_HELP_CONTACT = 903              '���ͷ���
Public Const ID_HELP_ONLINE = 904               '����ҽҵ
Public Const ID_HELP_ABOUT = 905                '����...

Public Const ID_PANE_PREVIEW = 10000
Public Const ID_PANE_INFO = 10001
Public Const ID_PANE_FILTER = 10002
Public Const ID_PANE_TEXTURE = 10003
