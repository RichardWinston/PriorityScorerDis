program CriticalThinkingScorerWrong5;

uses
  Vcl.Forms,
  frmMainUnit3 in 'frmMainUnit3.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
