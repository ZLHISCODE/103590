unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, OleCtrls, ZLDSVideoProcess_TLB, ComCtrls, ExtCtrls;

type
  TForm1 = class(TForm)
    TMCIAudio1: TTMCIAudio;
    butFormatCfg: TButton;
    Button2: TButton;
    edtPath: TEdit;
    Label1: TLabel;
    butStart: TButton;
    butPause: TButton;
    butStop: TButton;
    StatusBar1: TStatusBar;
    Timer1: TTimer;
    SaveDialog1: TSaveDialog;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    RadioButton5: TRadioButton;
    RadioButton6: TRadioButton;
    procedure Button2Click(Sender: TObject);
    procedure butFormatCfgClick(Sender: TObject);
    procedure butStartClick(Sender: TObject);
    procedure butPauseClick(Sender: TObject);
    procedure butStopClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button2Click(Sender: TObject);
begin
  SaveDialog1.FileName := edtPath.Text;
  if SaveDialog1.Execute then edtPath.Text := SaveDialog1.FileName;
end;

procedure TForm1.butFormatCfgClick(Sender: TObject);
begin
  TMCIAudio1.ShowFormatDialog;
end;

procedure TForm1.butStartClick(Sender: TObject);
var
  res: Boolean;
begin
  TMCIAudio1.RecordFile := edtPath.Text;

  res := TMCIAudio1.StartRecord;
  if not res then begin
    ShowMessage('录音程序启动失败。[' + TMCIAudio1.ErrorMsg + ']');
  end;
end;

procedure TForm1.butPauseClick(Sender: TObject);
begin
  TMCIAudio1.PauseRecord;
end;

procedure TForm1.butStopClick(Sender: TObject);
begin
  TMCIAudio1.StopRecord;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin

  if TMCIAudio1.RecordState = adsRun then begin
    StatusBar1.Panels[0].Text := '正在录音...';
  end else if TMCIAudio1.RecordState = adsPause then begin
    StatusBar1.Panels[0].Text := '暂停中...';
  end else begin
    StatusBar1.Panels[0].Text := '准备就绪...';
  end;

  if TMCIAudio1.RecordState <> adsStop then begin
    StatusBar1.Panels[1].Text := '录制长度：' + IntToStr(TMCIAudio1.RecordTimeLen) + '(秒)';
  end;
end;

procedure TForm1.RadioButton1Click(Sender: TObject);
begin
  TMCIAudio1.CompRate := TButton(Sender).Tag;
end;

end.
