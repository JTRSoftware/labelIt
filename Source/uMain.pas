Unit uMain;

{$mode Delphi}{$H+}

Interface

Uses
  //Lazarus
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Menus, ComCtrls, PrintersDlgs;

Type

  { TfMain }

  TfMain = Class(TForm)
    il16: TImageList;
    miInicioJanelas: TMenuItem;
    miFimJanelas: TMenuItem;
    miNovo: TMenuItem;
    miDisporMosaicoJanelas: TMenuItem;
    miDisporCascataJanelas: TMenuItem;
    miFecharTodasJanelas: TMenuItem;
    miMinimizarTodasJanelas: TMenuItem;
    odMain: TOpenDialog;
    miGuardarComoFicheiro: TMenuItem;
    miGuardarFicheiro: TMenuItem;
    miAbrirFicheiro: TMenuItem;
    mmMain: TMainMenu;
    sdMain: TSaveDialog;
    pdMain: TPrintDialog;
    Procedure miDisporCascataJanelasClick(ASender: TObject);
    Procedure miDisporMosaicoJanelasClick(ASender: TObject);
    Procedure miFecharTodasJanelasClick(ASender: TObject);
    Procedure miMinimizarTodasJanelasClick(ASender: TObject);
    Procedure miNovoClick(ASender: TObject);
    Procedure miAbrirFicheiroClick(ASender: TObject);
    Procedure miGuardarFicheiroClick(ASender: TObject);
    Procedure miGuardarComoFicheiroClick(ASender: TObject);
    Procedure SairClick(ASender: TObject);
  Private

  Public

  End;

Var
  fMain: TfMain;

Implementation

{$R *.lfm}

{ TfMain }

Uses
  //LabelIt
  uEtiqueta;

Procedure TfMain.miNovoClick(ASender: TObject);
Var
  LNewDoc: TfEtiqueta;
Begin
  LNewDoc := TfEtiqueta.Create(Application);
  LNewDoc.Caption := 'Etiqueta ' + IntToStr(MDIChildCount + 1);
  LNewDoc.Show;
End;

Procedure TfMain.miMinimizarTodasJanelasClick(ASender: TObject);
Var
  i: Integer;
Begin
  For i := 0 To Application.ComponentCount - 1 Do
    If (Application.Components[i] IS TForm) AND (Application.Components[i] <> Application.MainForm) Then
      (Application.Components[i] AS TForm).WindowState := wsMinimized;
End;

Procedure TfMain.miFecharTodasJanelasClick(ASender: TObject);
Var
  i: Integer;
Begin
  For i := 0 To Application.ComponentCount - 1 Do
    If (Application.Components[i] IS TForm) AND (Application.Components[i] <> Application.MainForm) Then
      (Application.Components[i] AS TForm).Close;
End;

Procedure TfMain.miDisporCascataJanelasClick(ASender: TObject);
Var
  i: Integer;
  LTop, LLeft: Integer;
Begin
  LTop := -32;
  LLeft := -32;

  For i := 0 To Application.ComponentCount - 1 Do
    If (Application.Components[i] IS TForm) AND (Application.Components[i] <> Application.MainForm) Then
    Begin
      TForm(Application.Components[i]).BringToFront;
      TForm(Application.Components[i]).Width := Round(Width * 0.6);
      TForm(Application.Components[i]).Height := Round(Height * 0.6);
      TForm(Application.Components[i]).Top := LTop;
      TForm(Application.Components[i]).Left := LLeft;
      LTop := LTop + 32;
      LLeft := LLeft + 32;
    End;
End;

Procedure TfMain.miDisporMosaicoJanelasClick(ASender: TObject);
Begin
  Tile;
End;

Procedure TfMain.miAbrirFicheiroClick(ASender: TObject);
Var
  LDoc: TfEtiqueta;
Begin
  If odMain.Execute Then
  Begin
    LDoc := TfEtiqueta.Create(Application);
    LDoc.EtiquetaFile.LoadFromFile(odMain.FileName);
    LDoc.LoadFromEtiquetaRecord;
    LDoc.Caption := ExtractFileName(odMain.FileName);
    LDoc.Show;
  End;
End;

Procedure TfMain.miGuardarFicheiroClick(ASender: TObject);
Begin
  If ActiveMDIChild IS TfEtiqueta Then
    TfEtiqueta(ActiveMDIChild).sbSaveFileClick(nil);
End;

Procedure TfMain.miGuardarComoFicheiroClick(ASender: TObject);
Begin
  If ActiveMDIChild IS TfEtiqueta Then
  Begin
    If sdMain.Execute Then
    Begin
      TfEtiqueta(ActiveMDIChild).EtiquetaFile.FileName := WideString(sdMain.FileName);
      TfEtiqueta(ActiveMDIChild).sbSaveFileClick(nil);
    End;
  End;
End;

Procedure TfMain.SairClick(ASender: TObject);
Begin
  Close;
End;

End.
