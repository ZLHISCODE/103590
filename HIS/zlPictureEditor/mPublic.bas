Attribute VB_Name = "mPublic"
Option Explicit
Public gfrmMain As fMain                '主窗体
Public gfDialogEx As fDialogEx          '
Public gfFilter As fFilter              '
Public gfOrientation As fOrientation    '
Public gfPanView As fPanView            '
Public gfPrint As fPrint                '
Public gfProperties As fProperties      '
Public gfResize As fResize              '
Public gfTexturize As fTexturize        '

Public gbln保留 As Boolean              '该图片对象是否保留，如果是，则不允许打开其他图片
Public glngSys As Long                  '系统号
