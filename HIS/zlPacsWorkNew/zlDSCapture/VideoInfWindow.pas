unit VideoInfWindow;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ComCtrls, VideoProcessDefine;

type
  TfrmVideoInf = class(TForm)
    RichEdit1: TRichEdit;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class function ShowVideoInf(parentHandle: HWND; videoInf: TVideoInf): WideString;
  end;


implementation

{$R *.dfm}

{ TfrmVideoInf }

class function TfrmVideoInf.ShowVideoInf(parentHandle: HWND; videoInf: TVideoInf): WideString;
var
  frmVideoInf: TfrmVideoInf;
begin
  try
    Result := '';

    if not Assigned(videoInf) then begin
      Result := '没有读取到当前视频的相关信息。';
      Exit;
    end;

    frmVideoInf := TfrmVideoInf.Create(Application{nil});
    try
      //frmVideoInf.ParentWindow := parentHandle;
      
      frmVideoInf.RichEdit1.Lines.Append('**************************************************');
      frmVideoInf.RichEdit1.Lines.Append('文件名：' + videoInf.videoFile);
      frmVideoInf.RichEdit1.Lines.Append('--------------------------------------------------');      
      frmVideoInf.RichEdit1.Lines.Append('视频格式：' + videoInf.MajorTypeName);
      frmVideoInf.RichEdit1.Lines.Append('编码格式：' + videoInf.SubTypeName);
      frmVideoInf.RichEdit1.Lines.Append('格式类型：' + videoInf.FormatTypeName);
      frmVideoInf.RichEdit1.Lines.Append('时间格式：' + videoInf.TimeFormatName);
      frmVideoInf.RichEdit1.Lines.Append('--------------------------------------------------');
      frmVideoInf.RichEdit1.Lines.Append('颜色深度：' + IntToStr(videoInf.VideoColorDepth));
      frmVideoInf.RichEdit1.Lines.Append('视频宽度：' + IntToStr(videoInf.VideoWidth));
      frmVideoInf.RichEdit1.Lines.Append('视频高度：' + IntToStr(videoInf.VideoHeight));
      frmVideoInf.RichEdit1.Lines.Append('--------------------------------------------------');
      frmVideoInf.RichEdit1.Lines.Append('流数量：' + IntToStr(videoInf.StreamCount));
      frmVideoInf.RichEdit1.Lines.Append('帧速率：' + FloatToStr(videoInf.FrameRate));
      frmVideoInf.RichEdit1.Lines.Append('帧数量：' + IntToStr(videoInf.FrameLen));      
      frmVideoInf.RichEdit1.Lines.Append('时长(秒)：' + IntToStr(videoInf.TimeLen));
      frmVideoInf.RichEdit1.Lines.Append('**************************************************');


      frmVideoInf.ShowModal();
    finally
      FreeAndNil(frmVideoInf);
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

procedure TfrmVideoInf.Button1Click(Sender: TObject);
begin
  Self.Close;
end;

end.
