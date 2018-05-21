program CriticalThinkingScorerDontCountWrong;

uses
  Vcl.Forms,
  frmMainUnit4 in 'frmMainUnit4.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
