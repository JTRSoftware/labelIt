Unit uDadosExternos;

{$mode Delphi}

Interface

Uses
  //LabelIt
  uLabelIt,
  //Lazarus
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ComCtrls, ExtCtrls, StdCtrls, Grids,
  //SQL
  MSSQLConn, SQLite3Conn, mysql80conn, mysql57conn, mysql56conn, mysql55conn, mysql51conn, mysql50conn, mysql41conn, mysql40conn, SQLDB,
  //SynEdit
  SynEdit, SynHighlighterSQL,
  //FPSpreadsheet
  fpspreadsheet;

Type

  { TfDadosExternos }

  TfDadosExternos = Class(TForm)
    bpgAnterior: TButton;
    bfcConcluir1: TButton;
    bqueryAnterior: TButton;
    bqueryConcluir: TButton;
    bfcSeguinte: TButton;
    bsqlSeguinte: TButton;
    bsqlAnterior: TButton;
    bliteAnterior: TButton;
    bfoAbrir1: TButton;
    bodSeguinte: TButton;
    bfcAnterior: TButton;
    bfoAbrir: TButton;
    bliteSeguinte: TButton;
    bsqlTest: TButton;
    bqueryTest: TButton;
    efcFicheiro: TEdit;
    eliteFicheiro: TEdit;
    gbsqlConfig: TGroupBox;
    lbpgPaginas: TListBox;
    lpgTitulo: TLabel;
    lsqlOutrosParams: TLabel;
    lliteFicheiro: TLabel;
    lfcTitulo: TLabel;
    lfcFicheiro: TLabel;
    lliteTitulo: TLabel;
    lodTitulo: TLabel;
    lsqlTitulo: TLabel;
    lqueryTitulo: TLabel;
    msqlOutrosParams: TMemo;
    MySQL40Con: TMySQL40Connection;
    MySQL41Con: TMySQL41Connection;
    MySQL50Con: TMySQL50Connection;
    MySQL51Con: TMySQL51Connection;
    MySQL55Con: TMySQL55Connection;
    MySQL56Con: TMySQL56Connection;
    MySQL57Con: TMySQL57Connection;
    MySQL80Con: TMySQL80Connection;
    odXLS: TOpenDialog;
    odLite: TOpenDialog;
    pfcNavegador1: TPanel;
    pliteNavegador1: TPanel;
    sgsqlProperties: TStringGrid;
    SQLServerCon: TMSSQLConnection;
    pcDadosExternos: TPageControl;
    psqlNavegador: TPanel;
    pliteNavegador: TPanel;
    podNavegador: TPanel;
    pfcNavegador: TPanel;
    rgOrigemDados: TRadioGroup;
    rgsqlTipo: TRadioGroup;
    SQLiteCon: TSQLite3Connection;
    SQLT: TSQLTransaction;
    seQuery: TSynEdit;
    SynSQL: TSynSQLSyn;
    tsPagina: TTabSheet;
    tsSQLQuery: TTabSheet;
    tsSQLServer: TTabSheet;
    tsSQLite: TTabSheet;
    tsFolhaCalculo: TTabSheet;
    tsOrigemDados: TTabSheet;
    Procedure FormShow(ASender: TObject);
    Procedure FormCreate(ASender: TObject);

    Procedure bfoAbrir1Click(ASender: TObject);
    Procedure bfoAbrirClick(ASender: TObject);
    Procedure bpgAnteriorClick(ASender: TObject);
    Procedure bfcAnteriorClick(ASender: TObject);
    Procedure bfcSeguinteClick(ASender: TObject);
    Procedure bqueryAnteriorClick(ASender: TObject);
    Procedure bliteAnteriorClick(ASender: TObject);
    Procedure bliteSeguinteClick(ASender: TObject);
    Procedure bodSeguinteClick(ASender: TObject);
    Procedure bsqlAnteriorClick(ASender: TObject);
    Procedure bsqlSeguinteClick(ASender: TObject);
    Procedure bfcConcluir1Click(ASender: TObject);
    Procedure bqueryConcluirClick(ASender: TObject);
    Procedure bsqlTestClick(ASender: TObject);
    Procedure bqueryTestClick(ASender: TObject);
    Procedure rgsqlTipoClick(ASender: TObject);

    Procedure sgsqlPropertiesSelectEditor(ASender: TObject; ACol, ARow: Integer; Var Editor: TWinControl);
    Procedure sgsqlPropertiesResize(ASender: TObject);
    Procedure tsSQLQueryShow(ASender: TObject);
  Private
    Function IsSelectQuery(Const AQuery: String): Boolean;
  Public
    pDadosExternos: ^TDadosExternos;
    Function GetDados(Var Dados: TDadosExternos): TModalResult;
    Procedure UpdateSQLProperties;
  End;

Var
  fDadosExternos: TfDadosExternos;

Implementation

{$R *.lfm}

{ TfDadosExternos }

Function TfDadosExternos.GetDados(Var Dados: TDadosExternos): TModalResult;
Var
  i: Integer;
Begin
  //Limpa dados
  For i := 0 To ComponentCount - 1 Do
  Begin
    If Components[i] IS TCustomEdit Then
      TCustomEdit(Components[i]).Text := ''
    Else If Components[i] IS TRadioGroup Then
      TRadioGroup(Components[i]).ItemIndex := 0
    Else If Components[i] IS TPageControl Then
      TPageControl(Components[i]).ActivePageIndex := 0;
  End;

  // Defaults
  rgOrigemDados.ItemIndex := 0;
  lbpgPaginas.Clear;
  msqlOutrosParams.Lines.Clear;
  seQuery.Lines.Clear;

  // Load existing data if any
  Case Dados.TipoDadosExternos Of
    tdeFolhaCalculo:
    Begin
      rgOrigemDados.ItemIndex := 0;
      efcFicheiro.Text := String(Dados.FicheiroDados);
      If (Dados.Params <> nil) AND (Length(Dados.Params) > 0) Then
      Begin
        // We can't easily select the sheet without loading the file first
        // But we can store it to select after loading if user clicks "Next"
      End;
    End;
    tdeSQLite:
    Begin
      rgOrigemDados.ItemIndex := 1;
      eliteFicheiro.Text := String(Dados.FicheiroDados);
    End;
    tdeSQLServer, tdeMySQL:
    Begin
      rgOrigemDados.ItemIndex := 2;
      If Dados.TipoDadosExternos = tdeSQLServer Then
        rgsqlTipo.ItemIndex := 0
      Else
        rgsqlTipo.ItemIndex := 1;

      // For now, let's just set the type
    End;
    tdeNenhum: rgOrigemDados.ItemIndex := 3;
  End;

  If (Dados.Query <> nil) AND (Length(Dados.Query) > 0) Then
  Begin
    seQuery.Lines.Clear;
    For i := 0 To High(Dados.Query) Do
      seQuery.Lines.Add(Dados.Query[i]);
  End;

  pDadosExternos := @Dados;
  pcDadosExternos.ActivePage := tsOrigemDados;

  Result := ShowModal;
End;

Procedure TfDadosExternos.FormCreate(ASender: TObject);
Begin
  pcDadosExternos.ShowTabs := False;
End;

Procedure TfDadosExternos.bodSeguinteClick(ASender: TObject);
Begin
  Case rgOrigemDados.ItemIndex Of
    0: pcDadosExternos.ActivePage := tsFolhaCalculo;
    1: pcDadosExternos.ActivePage := tsSQLite;
    2:
    Begin
      UpdateSQLProperties;
      pcDadosExternos.ActivePage := tsSQLServer;
    End;
    Else
    Begin
      pDadosExternos^.Clear;
      ModalResult := mrOk;
    End;
  End;
End;

Procedure TfDadosExternos.bsqlAnteriorClick(ASender: TObject);
Begin
  pcDadosExternos.ActivePage := tsOrigemDados;
End;

Procedure TfDadosExternos.bsqlSeguinteClick(ASender: TObject);
Begin
  //Verificar coerencia de dados
  lqueryTitulo.Caption := '... > Ligação servidor de base de dados > Query';
  pcDadosExternos.ActivePage := tsSQLQuery;
End;

Procedure TfDadosExternos.bfcAnteriorClick(ASender: TObject);
Begin
  pcDadosExternos.ActivePage := tsOrigemDados;
End;

Procedure TfDadosExternos.bfcSeguinteClick(ASender: TObject);
Var
  Workbook: TsWorkbook;
  SourceStream, DestStream: TFileStream;
  TempFileName: String;
  i: Integer;
Begin
  If efcFicheiro.Text <> '' Then
  Begin
    If FileExists(efcFicheiro.Text) Then
    Begin
      // Load Sheets
      lbpgPaginas.Clear;
      Workbook := TsWorkbook.Create;
      Try
        Try
          // Create a temp copy to avoid locking and ensure format detection works
          TempFileName := IncludeTrailingPathDelimiter(GetTempDir) + 'LabelIt_' + ExtractFileName(efcFicheiro.Text);

          SourceStream := TFileStream.Create(efcFicheiro.Text, fmOpenRead OR fmShareDenyNone);
          Try
            If SourceStream.Size = 0 Then
              Raise Exception.Create('O ficheiro está vazio ou bloqueado.');

            DestStream := TFileStream.Create(TempFileName, fmCreate);
            Try
              DestStream.CopyFrom(SourceStream, SourceStream.Size);
            Finally
              DestStream.Free;
            End;
          Finally
            SourceStream.Free;
          End;

          Try
            Workbook.ReadFromFile(TempFileName);
            For i := 0 To Workbook.GetWorksheetCount - 1 Do
            Begin
              lbpgPaginas.Items.Add(Workbook.GetWorksheetByIndex(i).Name);
            End;
          Finally
            If FileExists(TempFileName) Then DeleteFile(TempFileName);
          End;

        Except
          On E: Exception Do
          Begin
            ShowMessage('Erro ao importar dados: ' + E.Message + #13#10 + 'Ficheiro: ' + efcFicheiro.Text + #13#10 + 'Temp: ' + TempFileName);
            Exit;
          End;
        End;
      Finally
        Workbook.Free;
      End;

      If lbpgPaginas.Count > 0 Then
      Begin
        lbpgPaginas.ItemIndex := 0;
        // Try to select previously selected sheet if available
        If (pDadosExternos^.Params <> nil) AND (Length(pDadosExternos^.Params) > 0) Then
        Begin
          i := lbpgPaginas.Items.IndexOf(pDadosExternos^.Params[0]);
          If i >= 0 Then lbpgPaginas.ItemIndex := i;
        End;
      End;

      pcDadosExternos.ActivePage := tsPagina;
    End
    Else
      ShowMessage('Ficheiro não encontrado!');
  End
  Else
    ShowMessage('Ficheiro não definido!');
End;

Procedure TfDadosExternos.bpgAnteriorClick(ASender: TObject);
Begin
  pcDadosExternos.ActivePage := tsFolhaCalculo;
End;

Procedure TfDadosExternos.bfoAbrirClick(ASender: TObject);
Begin
  If odXLS.Execute Then
    efcFicheiro.Text := odXLS.FileName;
End;

Procedure TfDadosExternos.bfoAbrir1Click(ASender: TObject);
Begin
  If odLite.Execute Then
    eliteFicheiro.Text := odLite.FileName;
End;

Procedure TfDadosExternos.bqueryAnteriorClick(ASender: TObject);
Begin
  Case rgOrigemDados.ItemIndex Of
    1: pcDadosExternos.ActivePage := tsSQLite;
    2: pcDadosExternos.ActivePage := tsSQLServer;
    Else
      pcDadosExternos.ActivePage := tsOrigemDados;
  End;
End;

Procedure TfDadosExternos.bliteAnteriorClick(ASender: TObject);
Begin
  pcDadosExternos.ActivePage := tsOrigemDados;
End;

Procedure TfDadosExternos.bliteSeguinteClick(ASender: TObject);
Begin
  If eliteFicheiro.Text <> '' Then
  Begin
    If FileExists(eliteFicheiro.Text) Then
    Begin
      lqueryTitulo.Caption := '... > SQLite > Query';
      pcDadosExternos.ActivePage := tsSQLQuery;
    End
    Else
      ShowMessage('Ficheiro não encontrado!');
  End
  Else
    ShowMessage('Ficheiro não definido!');
End;

Procedure TfDadosExternos.bfcConcluir1Click(ASender: TObject);
Begin
  If lbpgPaginas.ItemIndex = -1 Then
  Begin
    ShowMessage('Selecione uma página/folha!');
    Exit;
  End;

  pDadosExternos^.TipoDadosExternos := tdeFolhaCalculo;
  pDadosExternos^.FicheiroDados := WideString(efcFicheiro.Text);
  SetLength(pDadosExternos^.Params, 1);
  pDadosExternos^.Params[0] := lbpgPaginas.Items[lbpgPaginas.ItemIndex];

  // Clear SQL specific fields
  pDadosExternos^.Servidor := '';
  pDadosExternos^.BaseDeDados := '';
  pDadosExternos^.Utilizador := '';
  pDadosExternos^.PalavraPasse := '';
  SetLength(pDadosExternos^.Query, 0);

  ModalResult := mrOk;
End;

Procedure TfDadosExternos.bqueryConcluirClick(ASender: TObject);
Var
  i: Integer;
  S: String;
Begin
  // Validate SQL data
  If seQuery.Lines.Count = 0 Then
  Begin
    ShowMessage('Defina a query SQL!');
    Exit;
  End;

  If NOT IsSelectQuery(seQuery.Text) Then
  Begin
    ShowMessage('Apenas são permitidas consultas de seleção (SELECT).');
    Exit;
  End;

  Case rgOrigemDados.ItemIndex Of
    1: pDadosExternos^.TipoDadosExternos := tdeSQLite;
    2:
    Begin
      If rgsqlTipo.ItemIndex = 0 Then
        pDadosExternos^.TipoDadosExternos := tdeSQLServer
      Else
        pDadosExternos^.TipoDadosExternos := tdeMySQL;
    End;
  End;

  pDadosExternos^.FicheiroDados := WideString(eliteFicheiro.Text); // For SQLite

  If rgOrigemDados.ItemIndex = 2 Then
  Begin
    pDadosExternos^.Servidor := sgsqlProperties.Cells[1, 1];
    pDadosExternos^.BaseDeDados := sgsqlProperties.Cells[1, 2];
    pDadosExternos^.Utilizador := sgsqlProperties.Cells[1, 3];
    pDadosExternos^.PalavraPasse := sgsqlProperties.Cells[1, 4];

    If rgsqlTipo.ItemIndex = 1 Then // MySQL
    Begin
      S := sgsqlProperties.Cells[1, 5];
      If S = '4.0' Then pDadosExternos^.Versao := msv40
      Else If S = '4.1' Then pDadosExternos^.Versao := msv41
      Else If S = '5.0' Then pDadosExternos^.Versao := msv50
      Else If S = '5.1' Then pDadosExternos^.Versao := msv51
      Else If S = '5.5' Then pDadosExternos^.Versao := msv55
      Else If S = '5.6' Then pDadosExternos^.Versao := msv56
      Else If S = '5.7' Then pDadosExternos^.Versao := msv57
      Else
        pDadosExternos^.Versao := msv80;
    End;
  End;

  SetLength(pDadosExternos^.Query, seQuery.Lines.Count);
  For i := 0 To seQuery.Lines.Count - 1 Do
    pDadosExternos^.Query[i] := seQuery.Lines[i];

  ModalResult := mrOk;
End;

Procedure TfDadosExternos.FormShow(ASender: TObject);
Begin
  Width := Round(Application.MainForm.Monitor.Width * 0.6);
  Height := Round(Application.MainForm.Monitor.Height * 0.6);
  Left := Application.MainForm.Monitor.Left + (Application.MainForm.Monitor.Width DIV 2) - (Width DIV 2);
  Top := Application.MainForm.Monitor.Top + (Application.MainForm.Monitor.Height DIV 2) - (Height DIV 2);
End;

Procedure TfDadosExternos.sgsqlPropertiesResize(ASender: TObject);
Begin
  sgsqlProperties.AutoSizeColumns;
  sgsqlProperties.ColWidths[1] := (sgsqlProperties.Width - 16) - sgsqlProperties.ColWidths[0];
End;

Procedure TfDadosExternos.tsSQLQueryShow(ASender: TObject);
Var
  Conn: TSQLConnection;
Begin
  Conn := nil;
  Case rgOrigemDados.ItemIndex Of
    1:
    Begin
      SynSQL.SQLDialect := sqlSQLite;
      SQLiteCon.Connected := False;
      SQLiteCon.DatabaseName := eliteFicheiro.Text;
      Conn := SQLiteCon;
    End;
    2:
    Begin
      Case rgsqlTipo.ItemIndex Of
        0:
        Begin
          SynSQL.SQLDialect := sqlMSSQL2K;
          SQLServerCon.Connected := False;
          SQLServerCon.HostName := sgsqlProperties.Cells[1, 1];
          SQLServerCon.DatabaseName := sgsqlProperties.Cells[1, 2];
          SQLServerCon.UserName := sgsqlProperties.Cells[1, 3];
          SQLServerCon.Password := sgsqlProperties.Cells[1, 4];
          Conn := SQLServerCon;
        End;
        1:
        Begin
          SynSQL.SQLDialect := sqlMySQL;
          Case pDadosExternos^.Versao Of
            msv40: Conn := MySQL40Con;
            msv41: Conn := MySQL41Con;
            msv50: Conn := MySQL50Con;
            msv51: Conn := MySQL51Con;
            msv55: Conn := MySQL55Con;
            msv56: Conn := MySQL56Con;
            msv57: Conn := MySQL57Con;
            msv80: Conn := MySQL80Con;
          End;
          If Assigned(Conn) Then
          Begin
            Conn.Connected := False;
            Conn.HostName := sgsqlProperties.Cells[1, 1];
            Conn.DatabaseName := sgsqlProperties.Cells[1, 2];
            Conn.UserName := sgsqlProperties.Cells[1, 3];
            Conn.Password := sgsqlProperties.Cells[1, 4];
          End;
        End;
      End;
    End;
    Else
      SynSQL.SQLDialect := sqlStandard;
  End;

  If Assigned(Conn) Then
  Begin
    Try
      Conn.Open;
      SynSQL.TableNames.Clear;
      Conn.GetTableNames(SynSQL.TableNames, False);
      Conn.Close;
    Except
      // Silently fail or log if connection fails here
    End;
  End;
End;

Procedure TfDadosExternos.bsqlTestClick(ASender: TObject);
Var
  Conn: TSQLConnection;
  S: String;
Begin
  Conn := nil;
  Try
    If rgOrigemDados.ItemIndex = 1 Then // SQLite
    Begin
      SQLiteCon.Connected := False;
      SQLiteCon.DatabaseName := eliteFicheiro.Text;
      Conn := SQLiteCon;
    End
    Else If rgOrigemDados.ItemIndex = 2 Then // SQL Server / MySQL
    Begin
      If rgsqlTipo.ItemIndex = 0 Then // SQL Server
      Begin
        SQLServerCon.Connected := False;
        SQLServerCon.HostName := sgsqlProperties.Cells[1, 1];
        SQLServerCon.DatabaseName := sgsqlProperties.Cells[1, 2];
        SQLServerCon.UserName := sgsqlProperties.Cells[1, 3];
        SQLServerCon.Password := sgsqlProperties.Cells[1, 4];
        Conn := SQLServerCon;
      End
      Else // MySQL
      Begin
        // Get version from grid if it exists
        If sgsqlProperties.RowCount > 5 Then
        Begin
          S := sgsqlProperties.Cells[1, 5];
          If S = '4.0' Then pDadosExternos^.Versao := msv40
          Else If S = '4.1' Then pDadosExternos^.Versao := msv41
          Else If S = '5.0' Then pDadosExternos^.Versao := msv50
          Else If S = '5.1' Then pDadosExternos^.Versao := msv51
          Else If S = '5.5' Then pDadosExternos^.Versao := msv55
          Else If S = '5.6' Then pDadosExternos^.Versao := msv56
          Else If S = '5.7' Then pDadosExternos^.Versao := msv57
          Else
            pDadosExternos^.Versao := msv80;
        End;

        Case pDadosExternos^.Versao Of
          msv40: Conn := MySQL40Con;
          msv41: Conn := MySQL41Con;
          msv50: Conn := MySQL50Con;
          msv51: Conn := MySQL51Con;
          msv55: Conn := MySQL55Con;
          msv56: Conn := MySQL56Con;
          msv57: Conn := MySQL57Con;
          msv80: Conn := MySQL80Con;
        End;
        If Assigned(Conn) Then
        Begin
          Conn.Connected := False;
          Conn.HostName := sgsqlProperties.Cells[1, 1];
          Conn.DatabaseName := sgsqlProperties.Cells[1, 2];
          Conn.UserName := sgsqlProperties.Cells[1, 3];
          Conn.Password := sgsqlProperties.Cells[1, 4];
        End;
      End;
    End;

    If Assigned(Conn) Then
    Begin
      Conn.Open;
      ShowMessage('Ligação estabelecida com sucesso!');
      Conn.Close;
    End
    Else
      ShowMessage('Configuração de ligação inválida.');
  Except
    On E: Exception Do
      ShowMessage('Erro na ligação: ' + E.Message);
  End;
End;

Procedure TfDadosExternos.bqueryTestClick(ASender: TObject);
Var
  Qry: TSQLQuery;
  Conn: TSQLConnection;
Begin
  If NOT IsSelectQuery(seQuery.Text) Then
  Begin
    ShowMessage('Apenas são permitidas consultas de seleção (SELECT).');
    Exit;
  End;

  Qry := TSQLQuery.Create(nil);
  Try
    Conn := nil;
    Case rgOrigemDados.ItemIndex Of
      1:
      Begin
        SQLiteCon.Connected := False;
        SQLiteCon.DatabaseName := eliteFicheiro.Text;
        Conn := SQLiteCon;
      End;
      2:
      Begin
        If rgsqlTipo.ItemIndex = 0 Then
        Begin
          SQLServerCon.Connected := False;
          SQLServerCon.HostName := sgsqlProperties.Cells[1, 1];
          SQLServerCon.DatabaseName := sgsqlProperties.Cells[1, 2];
          SQLServerCon.UserName := sgsqlProperties.Cells[1, 3];
          SQLServerCon.Password := sgsqlProperties.Cells[1, 4];
          Conn := SQLServerCon;
        End
        Else
        Begin
          Case pDadosExternos^.Versao Of
            msv40: Conn := MySQL40Con;
            msv41: Conn := MySQL41Con;
            msv50: Conn := MySQL50Con;
            msv51: Conn := MySQL51Con;
            msv55: Conn := MySQL55Con;
            msv56: Conn := MySQL56Con;
            msv57: Conn := MySQL57Con;
            msv80: Conn := MySQL80Con;
          End;
          If Assigned(Conn) Then
          Begin
            Conn.Connected := False;
            Conn.HostName := sgsqlProperties.Cells[1, 1];
            Conn.DatabaseName := sgsqlProperties.Cells[1, 2];
            Conn.UserName := sgsqlProperties.Cells[1, 3];
            Conn.Password := sgsqlProperties.Cells[1, 4];
          End;
        End;
      End;
    End;

    If Assigned(Conn) Then
    Begin
      SQLT.Database := Conn;
      Qry.Database := Conn;
      Qry.Transaction := SQLT;
      Qry.SQL.Text := seQuery.Text;
      Qry.Open;
      ShowMessage('Query executada com sucesso! Campos retornados: ' + IntToStr(Qry.FieldCount));
      Qry.Close;
    End;
  Finally
    Qry.Free;
  End;
End;

Procedure TfDadosExternos.UpdateSQLProperties;
Begin
  sgsqlProperties.ColCount := 2;
  sgsqlProperties.RowCount := 1;
  sgsqlProperties.Cells[0, 0] := 'Propriedade';
  sgsqlProperties.Cells[1, 0] := 'Valor';

  sgsqlProperties.RowCount := 5;
  sgsqlProperties.Cells[0, 1] := 'Servidor';
  sgsqlProperties.Cells[1, 1] := pDadosExternos^.Servidor;
  sgsqlProperties.Cells[0, 2] := 'Base de Dados';
  sgsqlProperties.Cells[1, 2] := pDadosExternos^.BaseDeDados;
  sgsqlProperties.Cells[0, 3] := 'Utilizador';
  sgsqlProperties.Cells[1, 3] := pDadosExternos^.Utilizador;
  sgsqlProperties.Cells[0, 4] := 'Palavra-Passe';
  sgsqlProperties.Cells[1, 4] := pDadosExternos^.PalavraPasse;

  If rgsqlTipo.ItemIndex = 1 Then // MySQL
  Begin
    sgsqlProperties.RowCount := 6;
    sgsqlProperties.Cells[0, 5] := 'Versão';
    Case pDadosExternos^.Versao Of
      msv40: sgsqlProperties.Cells[1, 5] := '4.0';
      msv41: sgsqlProperties.Cells[1, 5] := '4.1';
      msv50: sgsqlProperties.Cells[1, 5] := '5.0';
      msv51: sgsqlProperties.Cells[1, 5] := '5.1';
      msv55: sgsqlProperties.Cells[1, 5] := '5.5';
      msv56: sgsqlProperties.Cells[1, 5] := '5.6';
      msv57: sgsqlProperties.Cells[1, 5] := '5.7';
      msv80: sgsqlProperties.Cells[1, 5] := '8.0';
    End;
  End;

  sgsqlProperties.FixedCols := 1;
  sgsqlProperties.AutoSizeColumns;
  sgsqlProperties.ColWidths[1] := (sgsqlProperties.Width - 16) - sgsqlProperties.ColWidths[0];
End;

Procedure TfDadosExternos.rgsqlTipoClick(ASender: TObject);
Begin
  UpdateSQLProperties;
End;

Procedure TfDadosExternos.sgsqlPropertiesSelectEditor(ASender: TObject; ACol, ARow: Integer; Var Editor: TWinControl);
Begin
  If (ACol = 1) AND (ARow = 5) AND (rgsqlTipo.ItemIndex = 1) Then
  Begin
    Editor := sgsqlProperties.EditorByStyle(cbsPickList);
    If Assigned(Editor) Then
    Begin
      With TCustomComboBox(Editor) Do
      Begin
        Items.Clear;
        Items.Add('4.0');
        Items.Add('4.1');
        Items.Add('5.0');
        Items.Add('5.1');
        Items.Add('5.5');
        Items.Add('5.6');
        Items.Add('5.7');
        Items.Add('8.0');
      End;
    End;
  End;
End;

Function TfDadosExternos.IsSelectQuery(Const AQuery: String): Boolean;
Var
  S: String;
Begin
  S := Trim(AQuery);
  Result := (Length(S) >= 6) AND (UpperCase(Copy(S, 1, 6)) = 'SELECT');
End;

End.
