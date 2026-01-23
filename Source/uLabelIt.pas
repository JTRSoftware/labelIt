Unit uLabelIt;

{$mode Delphi}

Interface

Uses
  //LabelIt
  uLabelObjects,
  //Lazarus
  Classes, SysUtils, Forms, Graphics, Dialogs, System.UITypes;

Type
  TTamanho = Record
    Largura: Double;
    Altura: Double;
  End;

  TMargens = Record
    Cima: Double;
    Baixo: Double;
    Esquerda: Double;
    Direita: Double;
  End;

  TLayout = Record
    Cols: Integer;
    Rows: Integer;
    GapX: Double;
    GapY: Double;
  End;

  TObjeto = Record
    Objeto: TLabelItem;
  End;

  TTiposDadosExternos = (tdeNenhum, tdeFolhaCalculo, tdeSQLite, tdeSQLServer, tdeMySQL);
  TMySQLVersion = (msv40, msv41, msv50, msv51, msv55, msv56, msv57, msv80);

  TDadosExternos = Record
    TipoDadosExternos: TTiposDadosExternos;
    FicheiroDados: WideString;
    Versao: TMySQLVersion;
    Servidor: String;
    BaseDeDados: String;
    Utilizador: String;
    PalavraPasse: String;
    Params: TArray<String>;
    Query: TArray<String>;
    Procedure Clear;
  End;

  TModeloEtiqueta = Record
  Public
    Titulo: String;
    Impressora: String;
    Papel: String;
    Tamanho: TTamanho;
    Margens: TMargens;
    Orientacao: Integer; // 0=Portrait, 1=Landscape
    Layout: TLayout;
    Objetos: TArray<TObjeto>;
    Procedure Clear;
  End;

  TEtiqueta = Record
  Public
    FileName: WideString;
    Dados: TDadosExternos;
    Modelos: TArray<TModeloEtiqueta>;

    Procedure Clear;
    Procedure LoadFromFile(Const AFile: String);
    Procedure SaveToFile(Const AFile: String; AOverwrite: Boolean = True);
  End;

Implementation

Uses
  //Lazarus
  Zipper;

Procedure TModeloEtiqueta.Clear;
Begin
  Titulo := '';
  Impressora := '';
  Papel := '';
  Tamanho.Altura := 1;
  Tamanho.Largura := 1;
  Margens.Cima := 1;
  Margens.Baixo := 1;
  Margens.Esquerda := 1;
  Margens.Direita := 1;
  SetLength(Objetos, 0);
  Orientacao := 0;
  Layout.Cols := 1;
  Layout.Rows := 1;
  Layout.GapX := 0;
  Layout.GapY := 0;
End;

Procedure TEtiqueta.Clear;
Begin
  FileName := '';
  Dados.Clear;
  SetLength(Modelos, 0);
End;

Procedure TDadosExternos.Clear;
Begin
  TipoDadosExternos := tdeNenhum;
  FicheiroDados := '';
  Versao := msv80;
  Servidor := '';
  BaseDeDados := '';
  Utilizador := '';
  PalavraPasse := '';
  SetLength(Params, 0);
  SetLength(Query, 0);
End;

//Estrutura do ficheiro de dados (.lit\.zip)
{
<Ficheiro.lit>
  |_ E.0 (Pasta)
  |   |_ Main.r
  |   |_ Obj.0.r
  |   |_ Obj.1.r
  |   |_ ...
  |_ E.1 (Pasta)
  |   |_ Main.r
  |   |_ Obj.0.r
  |   |_ Obj.1.r
  |   |_ ...
  |_ SQL.r
  |  |_ Versao;
  |  |_ Servidor;
  |  |_ BaseDeDados;
  |  |_ Utilizador;
  |  |_ PalavraPasse;
  |  |_ Params
  |     |_ Linha[0]
  |     |_ Linha[1]
  |     |_ Linha[2]
  |     |_ ...
  |_ SQLQuery.sql (Query)
  |  |_ Linha[0]
  |  |_ Linha[1]
  |  |_ Linha[2]
  |  |_ ...
}

Procedure TEtiqueta.LoadFromFile(Const AFile: String);
Var
  LUnZipper: TUnZipper;
  LSL: TStringList;
  LTempDir: String;
  i, j: Integer;
  LObjFile: String;
  LMS: TMemoryStream;
  LReader: TReader;
  LModeloDir: String;
Begin
  If not FileExists(AFile) Then Exit;

  Clear;
  FileName := WideString(AFile);
  LTempDir := GetTempDir + 'LabelIt_' + IntToStr(GetTickCount64) + PathDelim;
  ForceDirectories(LTempDir);

  LUnZipper := TUnZipper.Create;
  LSL := TStringList.Create;
  Try
    LUnZipper.FileName := AFile;
    LUnZipper.OutputPath := LTempDir;
    LUnZipper.Examine;
    LUnZipper.UnZipAllFiles;

    // Load SQL.r
    If FileExists(LTempDir + 'SQL.r') Then
    Begin
      {$IFDEF FPC}
      LSL.LoadFromFile(LTempDir + 'SQL.r', TEncoding.UTF8);
      {$ELSE}
      LSL.LoadFromFile(LTempDir + 'SQL.r');
      {$ENDIF}
      Dados.TipoDadosExternos := TTiposDadosExternos(StrToIntDef(LSL.Values['TipoDadosExternos'], 0));
      Dados.FicheiroDados := LSL.Values['FicheiroDados'];
      Dados.Versao := TMySQLVersion(StrToIntDef(LSL.Values['Versao'], 8));
      Dados.Servidor := LSL.Values['Servidor'];
      Dados.BaseDeDados := LSL.Values['BaseDeDados'];
      Dados.Utilizador := LSL.Values['Utilizador'];
      Dados.PalavraPasse := LSL.Values['PalavraPasse'];
      
      // Deteção inteligente de ficheiros corrompidos
      If (Dados.TipoDadosExternos = tdeNenhum) Then
      Begin
        // Verificar se existem dados válidos que indicam corrupção
        If (Dados.FicheiroDados <> '') Then
        Begin
          // Ficheiro de folha de cálculo detetado
          If MessageDlg('Ficheiro corrompido',
                        'O ficheiro contém uma origem de dados (Folha de Cálculo) mas está marcado como "Sem Dados".' + sLineBreak +
                        'Isto pode indicar que o ficheiro está corrompido.' + sLineBreak + sLineBreak +
                        'Deseja corrigir automaticamente?',
                        mtWarning, [mbYes, mbNo], 0) = mrYes Then
            Dados.TipoDadosExternos := tdeFolhaCalculo;
        End
        Else If ((Dados.Servidor <> '') OR (Dados.BaseDeDados <> '')) AND
                FileExists(LTempDir + 'SQLQuery.sql') Then
        Begin
          // SQL Server detetado com query válida
          If MessageDlg('Ficheiro corrompido',
                        'O ficheiro contém uma origem de dados SQL Server mas está marcado como "Sem Dados".' + sLineBreak +
                        'Servidor: ' + Dados.Servidor + sLineBreak +
                        'Base de Dados: ' + Dados.BaseDeDados + sLineBreak + sLineBreak +
                        'Isto pode indicar que o ficheiro está corrompido.' + sLineBreak + sLineBreak +
                        'Deseja corrigir automaticamente?',
                        mtWarning, [mbYes, mbNo], 0) = mrYes Then
            Dados.TipoDadosExternos := tdeSQLServer;
        End;
      End;

      // Load Params
      SetLength(Dados.Params, 0);
      i := 0;
      While LSL.IndexOfName('Params[' + IntToStr(i) + ']') >= 0 Do
      Begin
        SetLength(Dados.Params, i + 1);
        Dados.Params[i] := LSL.Values['Params[' + IntToStr(i) + ']'];
        Inc(i);
      End;
    End;

    // Load SQLQuery.sql
    If FileExists(LTempDir + 'SQLQuery.sql') Then
    Begin
      {$IFDEF FPC}
      LSL.LoadFromFile(LTempDir + 'SQLQuery.sql', TEncoding.UTF8);
      {$ELSE}
      LSL.LoadFromFile(LTempDir + 'SQLQuery.sql');
      {$ENDIF}
      SetLength(Dados.Query, LSL.Count);
      For i := 0 To LSL.Count - 1 Do
        Dados.Query[i] := LSL[i];
    End;

    // Load Models
    i := 0;
    While DirectoryExists(LTempDir + 'E.' + IntToStr(i)) Do
    Begin
      LModeloDir := LTempDir + 'E.' + IntToStr(i) + PathDelim;
      SetLength(Modelos, i + 1);
      Modelos[i].Clear;

      // Load Main.r
      If FileExists(LModeloDir + 'Main.r') Then
      Begin
        {$IFDEF FPC}
        LSL.LoadFromFile(LModeloDir + 'Main.r', TEncoding.UTF8);
        {$ELSE}
        LSL.LoadFromFile(LModeloDir + 'Main.r');
        {$ENDIF}
        Modelos[i].Titulo := LSL.Values['Titulo'];
        If Modelos[i].Titulo = '' Then Modelos[i].Titulo := 'Etiqueta ' + IntToStr(i);
        Modelos[i].Impressora := LSL.Values['Impressora'];
        Modelos[i].Papel := LSL.Values['Papel'];
        Modelos[i].Tamanho.Largura := StrToFloatDef(LSL.Values['Tamanho.Largura'], 100);
        Modelos[i].Tamanho.Altura := StrToFloatDef(LSL.Values['Tamanho.Altura'], 50);
        Modelos[i].Margens.Cima := StrToFloatDef(LSL.Values['Margens.Cima'], 0);
        Modelos[i].Margens.Baixo := StrToFloatDef(LSL.Values['Margens.Baixo'], 0);
        Modelos[i].Margens.Esquerda := StrToFloatDef(LSL.Values['Margens.Esquerda'], 0);
        Modelos[i].Margens.Direita := StrToFloatDef(LSL.Values['Margens.Direita'], 0);
        Modelos[i].Orientacao := StrToIntDef(LSL.Values['Orientacao'], 0);
        Modelos[i].Layout.Cols := StrToIntDef(LSL.Values['Layout.Cols'], 1);
        Modelos[i].Layout.Rows := StrToIntDef(LSL.Values['Layout.Rows'], 1);
        Modelos[i].Layout.GapX := StrToFloatDef(LSL.Values['Layout.GapX'], 0);
        Modelos[i].Layout.GapY := StrToFloatDef(LSL.Values['Layout.GapY'], 0);
      End;

      // Load Objects
      j := 0;
      While FileExists(LModeloDir + 'Obj.' + IntToStr(j) + '.r') Do
      Begin
        LObjFile := LModeloDir + 'Obj.' + IntToStr(j) + '.r';
        LMS := TMemoryStream.Create;
        Try
          LMS.LoadFromFile(LObjFile);
          LReader := TReader.Create(LMS, 1024);
          Try
            SetLength(Modelos[i].Objetos, j + 1);
            Modelos[i].Objetos[j].Objeto := TLabelItem.CreateFromReader(LReader);
          Finally
            LReader.Free;
          End;
        Finally
          LMS.Free;
        End;
        Inc(j);
      End;
      Inc(i);
    End;

    // Compatibility check for old format (no E.x folders)
    If (i = 0) and FileExists(LTempDir + 'Main.r') Then
    Begin
      SetLength(Modelos, 1);
      Modelos[0].Clear;
      Modelos[0].Titulo := 'Etiqueta 0';
      {$IFDEF FPC}
      LSL.LoadFromFile(LTempDir + 'Main.r', TEncoding.UTF8);
      {$ELSE}
      LSL.LoadFromFile(LTempDir + 'Main.r');
      {$ENDIF}
      Modelos[0].Impressora := LSL.Values['Impressora'];
      Modelos[0].Papel := LSL.Values['Papel'];
      Modelos[0].Tamanho.Largura := StrToFloatDef(LSL.Values['Tamanho.Largura'], 100);
      Modelos[0].Tamanho.Altura := StrToFloatDef(LSL.Values['Tamanho.Altura'], 50);
      Modelos[0].Margens.Cima := StrToFloatDef(LSL.Values['Margens.Cima'], 0);
      Modelos[0].Margens.Baixo := StrToFloatDef(LSL.Values['Margens.Baixo'], 0);
      Modelos[0].Margens.Esquerda := StrToFloatDef(LSL.Values['Margens.Esquerda'], 0);
      Modelos[0].Margens.Direita := StrToFloatDef(LSL.Values['Margens.Direita'], 0);
      Modelos[0].Orientacao := StrToIntDef(LSL.Values['Orientacao'], 0);
      Modelos[0].Layout.Cols := StrToIntDef(LSL.Values['Layout.Cols'], 1);
      Modelos[0].Layout.Rows := StrToIntDef(LSL.Values['Layout.Rows'], 1);
      Modelos[0].Layout.GapX := StrToFloatDef(LSL.Values['Layout.GapX'], 0);
      Modelos[0].Layout.GapY := StrToFloatDef(LSL.Values['Layout.GapY'], 0);

      // Load Objects from root
      j := 0;
      While FileExists(LTempDir + 'Obj.' + IntToStr(j) + '.r') Do
      Begin
        LObjFile := LTempDir + 'Obj.' + IntToStr(j) + '.r';
        LMS := TMemoryStream.Create;
        Try
          LMS.LoadFromFile(LObjFile);
          LReader := TReader.Create(LMS, 1024);
          Try
            SetLength(Modelos[0].Objetos, j + 1);
            Modelos[0].Objetos[j].Objeto := TLabelItem.CreateFromReader(LReader);
          Finally
            LReader.Free;
          End;
        Finally
          LMS.Free;
        End;
        Inc(j);
      End;
    End;

  Finally
    LUnZipper.Free;
    LSL.Free;
  End;
End;

Procedure TEtiqueta.SaveToFile(Const AFile: String; AOverwrite: Boolean);
Var
  LZipper: TZipper;
  LSL: TStringList;
  LTempDir: String;
  i, j: Integer;
  LObjFile: String;
  LMS: TMemoryStream;
  LWriter: TWriter;
  LModeloDir: String;
Begin
  If FileExists(AFile) and (not AOverwrite) Then Exit;

  LTempDir := GetTempDir + 'LabelIt_Save_' + IntToStr(GetTickCount64) + PathDelim;
  ForceDirectories(LTempDir);

  LZipper := TZipper.Create;
  LSL := TStringList.Create;
  Try
    LZipper.FileName := AFile;

    // Save SQL.r
    LSL.Clear;
    LSL.Values['TipoDadosExternos'] := IntToStr(Ord(Dados.TipoDadosExternos));
    LSL.Values['FicheiroDados'] := Dados.FicheiroDados;
    LSL.Values['Versao'] := IntToStr(Ord(Dados.Versao));
    LSL.Values['Servidor'] := Dados.Servidor;
    LSL.Values['BaseDeDados'] := Dados.BaseDeDados;
    LSL.Values['Utilizador'] := Dados.Utilizador;
    LSL.Values['PalavraPasse'] := Dados.PalavraPasse;
    For i := 0 To High(Dados.Params) Do
      LSL.Values['Params[' + IntToStr(i) + ']'] := Dados.Params[i];
    {$IFDEF FPC}
    LSL.SaveToFile(LTempDir + 'SQL.r', TEncoding.UTF8);
    {$ELSE}
    LSL.SaveToFile(LTempDir + 'SQL.r');
    {$ENDIF}
    LZipper.Entries.AddFileEntry(LTempDir + 'SQL.r', 'SQL.r');

    // Save SQLQuery.sql
    LSL.Clear;
    For i := 0 To High(Dados.Query) Do
      LSL.Add(Dados.Query[i]);
    {$IFDEF FPC}
    LSL.SaveToFile(LTempDir + 'SQLQuery.sql', TEncoding.UTF8);
    {$ELSE}
    LSL.SaveToFile(LTempDir + 'SQLQuery.sql');
    {$ENDIF}
    LZipper.Entries.AddFileEntry(LTempDir + 'SQLQuery.sql', 'SQLQuery.sql');

    // Save Models
    For i := 0 To High(Modelos) Do
    Begin
      LModeloDir := LTempDir + 'E.' + IntToStr(i) + PathDelim;
      ForceDirectories(LModeloDir);

      // Save Main.r
      LSL.Clear;
      LSL.Values['Titulo'] := Modelos[i].Titulo;
      LSL.Values['Impressora'] := Modelos[i].Impressora;
      LSL.Values['Papel'] := Modelos[i].Papel;
      LSL.Values['Tamanho.Largura'] := FloatToStr(Modelos[i].Tamanho.Largura);
      LSL.Values['Tamanho.Altura'] := FloatToStr(Modelos[i].Tamanho.Altura);
      LSL.Values['Margens.Cima'] := FloatToStr(Modelos[i].Margens.Cima);
      LSL.Values['Margens.Baixo'] := FloatToStr(Modelos[i].Margens.Baixo);
      LSL.Values['Margens.Esquerda'] := FloatToStr(Modelos[i].Margens.Esquerda);
      LSL.Values['Margens.Direita'] := FloatToStr(Modelos[i].Margens.Direita);
      LSL.Values['Orientacao'] := IntToStr(Modelos[i].Orientacao);
      LSL.Values['Layout.Cols'] := IntToStr(Modelos[i].Layout.Cols);
      LSL.Values['Layout.Rows'] := IntToStr(Modelos[i].Layout.Rows);
      LSL.Values['Layout.GapX'] := FloatToStr(Modelos[i].Layout.GapX);
      LSL.Values['Layout.GapY'] := FloatToStr(Modelos[i].Layout.GapY);
      {$IFDEF FPC}
      LSL.SaveToFile(LModeloDir + 'Main.r', TEncoding.UTF8);
      {$ELSE}
      LSL.SaveToFile(LModeloDir + 'Main.r');
      {$ENDIF}
      LZipper.Entries.AddFileEntry(LModeloDir + 'Main.r', 'E.' + IntToStr(i) + '/Main.r');

      // Save Objects
      For j := 0 To High(Modelos[i].Objetos) Do
      Begin
        If Assigned(Modelos[i].Objetos[j].Objeto) Then
        Begin
          LObjFile := LModeloDir + 'Obj.' + IntToStr(j) + '.r';
          LMS := TMemoryStream.Create;
          Try
            LWriter := TWriter.Create(LMS, 1024);
            Try
              Modelos[i].Objetos[j].Objeto.SaveToWriter(LWriter);
            Finally
              LWriter.Free;
            End;
            LMS.SaveToFile(LObjFile);
            LZipper.Entries.AddFileEntry(LObjFile, 'E.' + IntToStr(i) + '/Obj.' + IntToStr(j) + '.r');
          Finally
            LMS.Free;
          End;
        End;
      End;
    End;

    LZipper.ZipAllFiles;

  Finally
    LZipper.Free;
    LSL.Free;
  End;
End;

End.
