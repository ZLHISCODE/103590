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
      Result := 'û�ж�ȡ����ǰ��Ƶ�������Ϣ��';
      Exit;
    end;

    frmVideoInf := TfrmVideoInf.Create(Application{nil});
    try
      //frmVideoInf.ParentWindow := parentHandle;
      
      frmVideoInf.RichEdit1.Lines.Append('**************************************************');
      frmVideoInf.RichEdit1.Lines.Append('�ļ�����' + videoInf.videoFile);
      frmVideoInf.RichEdit1.Lines.Append('--------------------------------------------------');      
      frmVideoInf.RichEdit1.Lines.Append('��Ƶ��ʽ��' + videoInf.MajorTypeName);
      frmVideoInf.RichEdit1.Lines.Append('�����ʽ��' + videoInf.SubTypeName);
      frmVideoInf.RichEdit1.Lines.Append('��ʽ���ͣ�' + videoInf.FormatTypeName);
      frmVideoInf.RichEdit1.Lines.Append('ʱ���ʽ��' + videoInf.TimeFormatName);
      frmVideoInf.RichEdit1.Lines.Append('--------------------------------------------------');
      frmVideoInf.RichEdit1.Lines.Append('��ɫ��ȣ�' + IntToStr(videoInf.VideoColorDepth));
      frmVideoInf.RichEdit1.Lines.Append('��Ƶ��ȣ�' + IntToStr(videoInf.VideoWidth));
      frmVideoInf.RichEdit1.Lines.Append('��Ƶ�߶ȣ�' + IntToStr(videoInf.VideoHeight));
      frmVideoInf.RichEdit1.Lines.Append('--------------------------------------------------');
      frmVideoInf.RichEdit1.Lines.Append('��������' + IntToStr(videoInf.StreamCount));
      frmVideoInf.RichEdit1.Lines.Append('֡���ʣ�' + FloatToStr(videoInf.FrameRate));
      frmVideoInf.RichEdit1.Lines.Append('֡������' + IntToStr(videoInf.FrameLen));      
      frmVideoInf.RichEdit1.Lines.Append('ʱ��(��)��' + IntToStr(videoInf.TimeLen));
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
