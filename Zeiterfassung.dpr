program Zeiterfassung;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {fMain},
  uSettings in 'uSettings.pas' {fSettings},
  uFunktionen in 'uFunktionen.pas',
  uDBFunktionen in 'uDBFunktionen.pas',
  uLegende in 'uLegende.pas' {fLegende},
  uExcelFunktionen in 'uExcelFunktionen.pas',
  uAbout in 'uAbout.pas' {fAbout};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfMain, fMain);
  Application.CreateForm(TfSettings, fSettings);
  Application.CreateForm(TfLegende, fLegende);
  Application.CreateForm(TfAbout, fAbout);
  Application.Run;
end.
