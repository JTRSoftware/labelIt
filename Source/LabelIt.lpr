Program LabelIt;

{$mode Delphi}{$H+}

Uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  {$IFDEF HASAMIGA}
  athreads,
  {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, runtimetypeinfocontrols,
  uMain,
  uEtiqueta,
  uLabelIt,
  uLabelObjects,
  uPrinting,
  uDadosExternos,
  uGridHelpers;

  {$R *.res}

Begin
  RequireDerivedFormResource := True;
  Application.Scaled:=True;
  {$PUSH}
  {$WARN 5044 OFF}
  Application.MainFormOnTaskbar := True;
  {$POP}
  Application.Initialize;
  Application.CreateForm(TfMain, fMain);
  Application.CreateForm(TfDadosExternos, fDadosExternos);
  Application.Run;
End.
