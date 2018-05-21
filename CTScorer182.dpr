program CTScorer182;

uses
  Vcl.Forms,
  frmMainUnit8 in 'frmMainUnit8.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
