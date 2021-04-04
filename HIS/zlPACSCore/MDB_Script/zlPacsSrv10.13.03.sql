Create Table DICOM胶片打印字体(
    影像类别 char(10),
    字体大小 char(10),
    是否随图像缩放 yesno,
    PRIMARY KEY (影像类别));

update 版本表 set 版本号='10.13.03';