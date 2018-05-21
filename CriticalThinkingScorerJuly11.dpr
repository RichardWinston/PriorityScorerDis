program CriticalThinkingScorerJuly11;

uses
  Vcl.Forms,
  frmMainUnit7 in 'frmMainUnit7.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
