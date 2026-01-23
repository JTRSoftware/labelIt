Unit uEtiqueta;

{$mode Delphi}{$H+}

Interface

Uses
  //LabelIt
  uLabelObjects, uLabelIt, MSSQLConn, mysql80conn, mysql57conn, mysql56conn, mysql55conn, mysql51conn, mysql50conn, mysql41conn, mysql40conn, SQLite3Conn, SQLDB,
  //FPSpreadsheet
  fpspreadsheet, fpSpreadsheetCtrls, fpsTypes,
  //Lazarus
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Grids, ColorBox, Menus, LCLType, LCLIntf, Types, ComCtrls, Printers,
  Spin, Buttons, StdCtrls, ExtDlgs, RTTICtrls, Contnrs;

Type
  TTabSettings = Class
  Public
    LabelManager: TLabelManager;
    PrinterName: String;
    PaperName: String;
    Margins: TRect;
    Orientation: TPrinterOrientation;
    LayoutCols: Integer;
    LayoutRows: Integer;
    LayoutGapX: Integer;
    LayoutGapY: Integer;
    PaperWidth: Integer;
    PaperHeight: Integer;
    Constructor Create;
    Destructor Destroy; Override;
  End;

  TResizeHandle = (rhNone, rhTopLeft, rhTop, rhTopRight, rhRight, rhBottomRight, rhBottom, rhBottomLeft, rhLeft);

  { TfEtiqueta }

  TfEtiqueta = Class(TForm)
    bPrint: TBitBtn;
    cbAlignTo: TComboBox;
    lDataStaus: TLabel;
    miTabRename: TMenuItem;
    miTabDelete: TMenuItem;
    miDataCheckAll: TMenuItem;
    miDataUncheckAll: TMenuItem;
    miDeleteCol: TMenuItem;
    miDeleteRow: TMenuItem;
    miGridDelete: TMenuItem;
    miInsertColRight: TMenuItem;
    miInsertRowAbove: TMenuItem;
    miInsertRowBelow: TMenuItem;
    miInsertColLeft: TMenuItem;
    miInsert: TMenuItem;
    miUngroup: TMenuItem;
    miGroup: TMenuItem;
    miColar: TMenuItem;
    miCortar: TMenuItem;
    miCopiar: TMenuItem;
    MySQL40Con: TMySQL40Connection;
    MySQL41Con: TMySQL41Connection;
    MySQL50Con: TMySQL50Connection;
    MySQL51Con: TMySQL51Connection;
    MySQL55Con: TMySQL55Connection;
    MySQL56Con: TMySQL56Connection;
    MySQL57Con: TMySQL57Connection;
    MySQL80Con: TMySQL80Connection;
    opdImagem: TOpenPictureDialog;
    pcEtiquetas: TPageControl;
    pnlData: TPanel;
    pmData: TPopupMenu;
    pmTab: TPopupMenu;
    pSeparador2: TPanel;
    pSeparador3: TPanel;
    pSeparador4: TPanel;
    pSeparador5: TPanel;
    pTools: TPanel;
    pDataTools: TPanel;
    sbAutoSizeCols: TSpeedButton;
    sbEqualizeWidth: TSpeedButton;
    sbEqualizeHeight: TSpeedButton;
    sbRefreshData: TSpeedButton;
    sbSaveFile: TSpeedButton;
    sbExternalData: TSpeedButton;
    sbNewGrid: TSpeedButton;
    sbNewBarCode: TSpeedButton;
    sbNewImage: TSpeedButton;
    sbNewText: TSpeedButton;
    Separator1: TMenuItem;
    sePrintCount: TSpinEdit;
    sbAlignObjectCenterHorizontal: TSpeedButton;
    sbAlignObjectCenterVertical: TSpeedButton;
    sbAlignObjectBottom: TSpeedButton;
    sbAlignObjectTop: TSpeedButton;
    sbAlignObjectLeft: TSpeedButton;
    sbAlignObjectRight: TSpeedButton;
    sbExpandMargins: TSpeedButton;
    sgData: TStringGrid;
    SplitterData: TSplitter;
    pmContext: TPopupMenu;
    miDelete: TMenuItem;
    miDuplicate: TMenuItem;
    miSeparator1: TMenuItem;
    miBringToFront: TMenuItem;
    miSendToBack: TMenuItem;
    miSeparator2: TMenuItem;
    miProperties: TMenuItem;
    Splitter2: TSplitter;
    SQLiteCon: TSQLite3Connection;
    SQLServerCon: TMSSQLConnection;
    SQLT: TSQLTransaction;
    tsMais: TTabSheet;
    tsEtiqueta0: TTabSheet;
    sbContainer0: TScrollBox;
    pbDesign0: TPaintBox;
    pObjetPlus0: TPanel;
    tvObjetos0: TTreeView;
    sgProperties0: TStringGrid;
    splitterTVSG0: TSplitter;
    cbColorEditor0: TColorBox;
    tmrFilter: TTimer;
    Qry: TSQLQuery;
    Procedure FormClose(ASender: TObject; Var ACloseAction: TCloseAction);
    Procedure FormCreate(ASender: TObject);
    Procedure FormDestroy(ASender: TObject);

    Procedure pbDesignPaint(ASender: TObject);
    Procedure pbDesignMouseDown(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);
    Procedure pbDesignMouseMove(ASender: TObject; AShift: TShiftState; AX, AY: Integer);
    Procedure pbDesignMouseUp(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);
    Procedure pObjetPlusResize(ASender: TObject);
    Procedure sbAutoSizeColsClick(ASender: TObject);
    Procedure sbExternalDataClick(ASender: TObject);
    Procedure sgPropertiesSelectCell(ASender: TObject; ACol, ARow: Integer; Var ACanSelect: Boolean);
    Procedure sgPropertiesSetEditText(ASender: TObject; ACol, ARow: Integer; Const AValue: String);
    Procedure sgPropertiesMouseDown(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);
    Procedure sgPropertiesDrawCell(ASender: TObject; ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState);
    Procedure sgPropertiesSelectEditor(ASender: TObject; ACol, ARow: Integer; Var AEditor: TWinControl);
    Procedure sgPropertiesButtonClick(ASender: TObject; ACol, ARow: Integer);
    Procedure cbColorEditorChange(ASender: TObject);
    Procedure cbColorEditorExit(ASender: TObject);

    Procedure miDeleteTabClick(ASender: TObject);
    Procedure DeleteActiveTab;

    Procedure miDeleteClick(ASender: TObject);
    Procedure miDuplicateClick(ASender: TObject);
    Procedure miBringToFrontClick(ASender: TObject);
    Procedure miSendToBackClick(ASender: TObject);
    Procedure miPropertiesClick(ASender: TObject);
    Procedure miInsertColLeftClick(ASender: TObject);
    Procedure miInsertColRightClick(ASender: TObject);
    Procedure miInsertRowAboveClick(ASender: TObject);
    Procedure miInsertRowBelowClick(ASender: TObject);
    Procedure miDeleteColClick(ASender: TObject);
    Procedure miDeleteRowClick(ASender: TObject);

    Procedure pmContextPopup(ASender: TObject);

    Procedure PasteFromClipboard(ASender: TObject);
    Procedure CopyToClipboard(ASender: TObject);
    Procedure CutToClipboard(ASender: TObject);

    Procedure miGroupClick(ASender: TObject);
    Procedure miUngroupClick(ASender: TObject);

    Procedure sbNewBarCodeClick(ASender: TObject);
    Procedure sbNewImageClick(ASender: TObject);
    Procedure sbNewTextClick(ASender: TObject);
    Procedure sbNewGridClick(ASender: TObject);
    Procedure sbSaveFileClick(ASender: TObject);
    Procedure UpdateEtiquetaRecord;
    Procedure LoadFromEtiquetaRecord;

    Procedure tvObjetosChange(ASender: TObject; ANode: TTreeNode);
    Procedure tvObjetosSelectionChanged(ASender: TObject);

    Procedure sbAlignObjectClick(ASender: TObject);
    Procedure sbExpandMarginsClick(ASender: TObject);
    Procedure sbEqualizeWidthClick(ASender: TObject);
    Procedure sbEqualizeHeightClick(ASender: TObject);
    Procedure sbContainerMouseDown(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);

    Procedure sgPropertiesEditorKeyPress(ASender: TObject; Var AKey: Char);

    Procedure sbRefreshDataClick(ASender: TObject);
    Procedure miDataCheckAllClick(ASender: TObject);
    Procedure miDataUncheckAllClick(ASender: TObject);

    Procedure tmrFilterTimer(ASender: TObject);
    Procedure sgDataSelectCell(ASender: TObject; ACol, ARow: Integer; Var ACanSelect: Boolean);
    Procedure sgDataDrawCell(ASender: TObject; ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState);
    Procedure sgDataPrepareCanvas(ASender: TObject; ACol, ARow: Integer; AState: TGridDrawState);
    Procedure sgDataSetEditText(ASender: TObject; ACol, ARow: Integer; Const AValue: String);
    Procedure sgDataMouseDown(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);
    Procedure pcEtiquetasChange(ASender: TObject);
  Private
    FDragItem: TLabelItem;
    FDragStartPos: TPoint;
    FIsDragging: Boolean;
    FSelectedItem: TLabelItem;
    FIsResizing: Boolean;
    FResizeHandle: TResizeHandle;
    FResizeStartRect: TRect;
    FIsResizingGridLine: Boolean;
    FResizingGrid: TLabelGrid;
    FResizingLineIndex: Integer;
    FResizingIsRow: Boolean;
    FMarginsExpanded: Boolean;
    FBordersExpanded: Boolean;
    FCellPropsExpanded: Boolean;
    FGridPropsExpanded: Boolean;
    FMosaicoExpanded: Boolean;
    FFontExpanded: Boolean;
    FFontStyleExpanded: Boolean;
    FUpdatingTree: Boolean;
    FClipboard: TLabelItem; // Clipboard storage
    FDataLoadedOnce: Boolean;
    FLastFilterValue: String;
    FIsAddingTab: Boolean;

    Function GetActiveSettings: TTabSettings;
    Function GetActiveLabelManager: TLabelManager;
    Function GetActivePaintBox: TPaintBox;
    Function GetActiveScrollBox: TScrollBox;
    Function GetActiveTreeView: TTreeView;
    Function GetActivePropertiesGrid: TStringGrid;

    Procedure DeleteSelectedItem;

    Procedure UpdatePropertiesPanel(AItem: TLabelItem);
    Procedure ShowLabelProperties;
    Procedure DrawSelectionHandles(ACanvas: TCanvas; AItem: TLabelItem);
    Function DetectHandle(AItem: TLabelItem; AX, AY: Integer): TResizeHandle;
    Procedure UpdateLabelSize;
    Procedure RebuildObjectTree;
    Procedure SelectInTree(AItem: TLabelItem);
    Function IsPaperSizeEditable(Const APaperName: String): Boolean;
    Function IsPaperSizeContinuous(Const APaperName: String): Boolean;
    Function PxToMm(APx: Integer): Double;
    Function MmToPx(AMm: Double): Integer;
    Procedure LoadExternalData;
    Procedure UpdateObjectsFromData(ARow: Integer = -1);
    Procedure AddNewTab(Const ATitulo: String = '');
  Private
    FDataLoader: TThread;
    Procedure StopDataLoader;
  Public
    EtiquetaFile: TEtiqueta;
    Property LabelManager: TLabelManager Read GetActiveLabelManager;
  End;

  { TDataLoaderThread }
  TDataLoaderThread = Class(TThread)
  Private
    FOwner: TForm;
    FDados: TDadosExternos;
    FGrid: TStringGrid;
    FHeaders: TStringList;
    FBatchRows: TList; // List of TStringList
    FIsFinished: Boolean;
    FFilters: TStringList; // NomeColuna=ValorFiltro
    FCheckedKeys: TStringList; // Conteúdo das linhas que estavam marcadas
    FError: String;
    FPerformAutoSize: Boolean;
    Procedure SyncBatch;
    Procedure SyncError;
  Protected
    Procedure Execute; Override;
  Public
    Constructor Create(AOwner: TForm; ADados: TDadosExternos; AGrid: TStringGrid; AFilters: TStringList; ACheckedKeys: TStringList; APerformAutoSize: Boolean);
    Destructor Destroy; Override;
  End;

Var
  fEtiqueta: TfEtiqueta;

Const
  sMenos = #$E2#$9E#$96; // UTF-8 for ➖
  sMais = #$E2#$9E#$95;  // UTF-8 for ➕
  sPropIndent = '    ';
  iCheckColWidth = 24;

Implementation

{$R *.lfm}

{ TfEtiqueta }

Uses
  //Labelit
  uDadosExternos, uMain, uGridHelpers;

{ TTabSettings }

Constructor TTabSettings.Create;
Begin
  LabelManager := TLabelManager.Create;
  PrinterName := '';
  PaperName := '';
  Margins := Rect(10, 10, 10, 10);
  Orientation := poPortrait;
  LayoutCols := 1;
  LayoutRows := 1;
  LayoutGapX := 0;
  LayoutGapY := 0;
  PaperWidth := 0;
  PaperHeight := 0;
End;

Destructor TTabSettings.Destroy;
Begin
  LabelManager.Free;
  Inherited Destroy;
End;

Procedure TfEtiqueta.FormCreate(ASender: TObject);
Var
  LSettings: TTabSettings;
Begin
  FDataLoadedOnce := False;
  FIsAddingTab := False;

  // tsEtiqueta0 is our template tab, keep it hidden and initialize it
  LSettings := TTabSettings.Create;
  tsEtiqueta0.Tag := PtrInt(LSettings);
  tsEtiqueta0.TabVisible := False;

  // Initialize Defaults from Printer for the template tab
  If Printer.Printers.Count > 0 Then
  Begin
    LSettings.PrinterName := Printer.Printers[Printer.PrinterIndex];
    Try
      LSettings.PaperName := Printer.PaperSize.PaperName;
    Except
      LSettings.PaperName := '';
    End;
  End;

  // Por omissão, todas as secções começam fechadas
  FMarginsExpanded := False;
  FBordersExpanded := False;
  FCellPropsExpanded := False;
  FGridPropsExpanded := False;
  FMosaicoExpanded := False;

  // Enable form to capture keyboard events for shortcuts
  KeyPreview := True;

  // Add the first visible tab dynamically
  AddNewTab('Etiqueta 0');
End;

Procedure TfEtiqueta.FormClose(ASender: TObject; Var ACloseAction: TCloseAction);
Begin
  Self.Release;
End;

Procedure TfEtiqueta.FormDestroy(ASender: TObject);
Var
  i: Integer;
Begin
  // Free settings for all tabs
  For i := 0 To pcEtiquetas.PageCount - 1 Do
    If pcEtiquetas.Pages[i].Tag <> 0 Then
      TTabSettings(pcEtiquetas.Pages[i].Tag).Free;

  StopDataLoader;
  If Assigned(FClipboard) Then FClipboard.Free;
End;

Function TfEtiqueta.GetActiveSettings: TTabSettings;
Begin
  Result := nil;
  If (pcEtiquetas.ActivePage <> nil) AND (pcEtiquetas.ActivePage.Tag <> 0) Then
    Result := TTabSettings(pcEtiquetas.ActivePage.Tag);
End;

Function TfEtiqueta.GetActiveLabelManager: TLabelManager;
Var
  LSettings: TTabSettings;
Begin
  LSettings := GetActiveSettings;
  If LSettings <> nil Then
    Result := LSettings.LabelManager
  Else
    Result := nil;
End;

Function TfEtiqueta.GetActivePaintBox: TPaintBox;
Begin
  If pcEtiquetas.ActivePage = tsEtiqueta0 Then Exit(pbDesign0);
  Result := nil;
  If pcEtiquetas.ActivePage <> nil Then
    Result := TPaintBox(pcEtiquetas.ActivePage.FindComponent('pbDesign'));
End;

Function TfEtiqueta.GetActiveScrollBox: TScrollBox;
Begin
  If pcEtiquetas.ActivePage = tsEtiqueta0 Then Exit(sbContainer0);
  Result := nil;
  If pcEtiquetas.ActivePage <> nil Then
    Result := TScrollBox(pcEtiquetas.ActivePage.FindComponent('sbContainer'));
End;

Function TfEtiqueta.GetActiveTreeView: TTreeView;
Begin
  If pcEtiquetas.ActivePage = tsEtiqueta0 Then Exit(tvObjetos0);
  Result := nil;
  If pcEtiquetas.ActivePage <> nil Then
    Result := TTreeView(pcEtiquetas.ActivePage.FindComponent('tvObjetos'));
End;

Function TfEtiqueta.GetActivePropertiesGrid: TStringGrid;
Begin
  If pcEtiquetas.ActivePage = tsEtiqueta0 Then Exit(sgProperties0);
  Result := nil;
  If pcEtiquetas.ActivePage <> nil Then
    Result := TStringGrid(pcEtiquetas.ActivePage.FindComponent('sgProperties'));
End;

Procedure TfEtiqueta.AddNewTab(Const ATitulo: String);
Var
  LTab: TTabSheet;
  LSettings: TTabSettings;
  LSB: TScrollBox;
  LPB: TPaintBox;
  LTV: TTreeView;
  LSG: TStringGrid;
  LPanel: TPanel;
  LSplitter: TSplitter;
Begin
  LTab := TTabSheet.Create(pcEtiquetas);
  LTab.PageControl := pcEtiquetas;
  LTab.Caption := ATitulo;
  If LTab.Caption = '' Then LTab.Caption := 'Etiqueta ' + IntToStr(pcEtiquetas.PageCount - 2);
  
  // Move before tsMais
  LTab.PageIndex := pcEtiquetas.PageCount - 2;

  LSettings := TTabSettings.Create;
  LTab.Tag := PtrInt(LSettings);

  // Initialize Defaults from Printer
  If Printer.Printers.Count > 0 Then
  Begin
    LSettings.PrinterName := Printer.Printers[Printer.PrinterIndex];
    Try
      LSettings.PaperName := Printer.PaperSize.PaperName;
    Except
      LSettings.PaperName := '';
    End;
  End;

  // Create Components (Cloning template structure)
  // Owner is LTab for automatic memory management
  // ScrollBox
  LSB := TScrollBox.Create(LTab);
  LSB.Name := 'sbContainer';
  LSB.Parent := LTab;
  LSB.Align := alClient;
  LSB.Color := clGray;
  LSB.ParentBackground := False;
  LSB.ParentColor := False;
  LSB.OnMouseDown := sbContainerMouseDown;
  LSB.PopupMenu := pmContext;

  // PaintBox
  LPB := TPaintBox.Create(LTab);
  LPB.Name := 'pbDesign';
  LPB.Parent := LSB;
  LPB.Left := 8;
  LPB.Top := 8;
  LPB.Color := clWhite;
  LPB.ParentColor := False;
  LPB.OnPaint := pbDesignPaint;
  LPB.OnMouseDown := pbDesignMouseDown;
  LPB.OnMouseMove := pbDesignMouseMove;
  LPB.OnMouseUp := pbDesignMouseUp;
  LPB.PopupMenu := pmContext;

  // Right Panel for Tree and Properties
  LPanel := TPanel.Create(LTab);
  LPanel.Name := 'pObjetPlus';
  LPanel.Parent := LTab;
  LPanel.Width := 350;
  LPanel.Align := alRight;
  LPanel.OnResize := pObjetPlusResize;

  // TreeView
  LTV := TTreeView.Create(LTab);
  LTV.Name := 'tvObjetos';
  LTV.Parent := LPanel;
  LTV.Align := alTop;
  LTV.Top := 0;
  LTV.Height := 150;
  LTV.ReadOnly := True;
  LTV.HideSelection := False;
  LTV.OnChange := tvObjetosChange;
  LTV.OnSelectionChanged := tvObjetosSelectionChanged;

  // Splitter
  LSplitter := TSplitter.Create(LTab);
  LSplitter.Name := 'splitterTVSG';
  LSplitter.Parent := LPanel;
  LSplitter.Align := alTop;
  LSplitter.Top := 1000;
  LSplitter.Cursor := crVSplit;

  // Properties Grid
  LSG := TStringGrid.Create(LTab);
  LSG.Name := 'sgProperties';
  LSG.Parent := LPanel;
  LSG.Align := alClient;
  LSG.ColCount := 2;
  LSG.FixedCols := 0;
  LSG.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing, goSmoothScroll];
  
  // Setup Columns
  LSG.Columns.Add.Title.Caption := 'Propriedade';
  LSG.Columns[0].Title.Font.Style := [fsBold];
  LSG.Columns[0].Width := 120;
  LSG.Columns.Add.Title.Caption := 'Valor';
  LSG.Columns[1].Title.Font.Style := [fsBold];
  
  LSG.OnDrawCell := sgPropertiesDrawCell;
  LSG.OnMouseDown := sgPropertiesMouseDown;
  LSG.OnSelectEditor := sgPropertiesSelectEditor;
  LSG.OnSelectCell := sgPropertiesSelectCell;
  LSG.OnSetEditText := sgPropertiesSetEditText;

  pcEtiquetas.ActivePage := LTab;
  tsMais.PageIndex := pcEtiquetas.PageCount - 1;
  
  UpdateLabelSize;
  RebuildObjectTree;
  ShowLabelProperties;
  
  // Force initial resize for the properties grid
  pObjetPlusResize(LPanel);
End;

Procedure TfEtiqueta.StopDataLoader;
Begin
  If Assigned(FDataLoader) Then
  Begin
    FDataLoader.Terminate;
    FDataLoader.WaitFor;
    FreeAndNil(FDataLoader);
  End;
End;

Procedure TfEtiqueta.ShowLabelProperties;
Var
  LSG: TStringGrid;
  LSettings: TTabSettings;

  Procedure AddProp(Const AName, AValue: String);
  Var
    LRow: Integer;
  Begin
    LRow := LSG.RowCount;
    LSG.RowCount := LRow + 1;
    LSG.Cells[0, LRow] := AName;
    LSG.Cells[1, LRow] := AValue;
  End;

Begin
  LSettings := GetActiveSettings;
  If LSettings = nil Then Exit;

  LSG := GetActivePropertiesGrid;
  If LSG = nil Then Exit;

  LSG.Enabled := True;
  LSG.RowCount := 1; // Reset

  AddProp('Impressora', LSettings.PrinterName);
  AddProp('Papel', LSettings.PaperName);

  If FMarginsExpanded Then
    AddProp(sMenos + 'Margens', '')
  Else
    AddProp(sMais + 'Margens', '');

  If FMarginsExpanded Then
  Begin
    AddProp(sPropIndent + 'Esquerda (mm)', FloatToStrF(PxToMm(LSettings.Margins.Left), ffFixed, 7, 2));
    AddProp(sPropIndent + 'Topo (mm)', FloatToStrF(PxToMm(LSettings.Margins.Top), ffFixed, 7, 2));
    AddProp(sPropIndent + 'Direita (mm)', FloatToStrF(PxToMm(LSettings.Margins.Right), ffFixed, 7, 2));
    AddProp(sPropIndent + 'Baixo (mm)', FloatToStrF(PxToMm(LSettings.Margins.Bottom), ffFixed, 7, 2));
  End;

  If LSettings.Orientation = poLandscape Then
    AddProp('Orientação', 'Landscape')
  Else
    AddProp('Orientação', 'Portrait');

  If FMosaicoExpanded Then
    AddProp(sMenos + 'Mosaico', '')
  Else
    AddProp(sMais + 'Mosaico', '');

  If FMosaicoExpanded Then
  Begin
    AddProp(sPropIndent + 'Colunas', IntToStr(LSettings.LayoutCols));
    AddProp(sPropIndent + 'Linhas', IntToStr(LSettings.LayoutRows));
    AddProp(sPropIndent + 'Espaço X (mm)', FloatToStrF(PxToMm(LSettings.LayoutGapX), ffFixed, 7, 2));
    AddProp(sPropIndent + 'Espaço Y (mm)', FloatToStrF(PxToMm(LSettings.LayoutGapY), ffFixed, 7, 2));
  End;

  AddProp('Largura (mm)', FloatToStrF(PxToMm(LSettings.PaperWidth), ffFixed, 7, 2));
  AddProp('Altura (mm)', FloatToStrF(PxToMm(LSettings.PaperHeight), ffFixed, 7, 2));

  pObjetPlusResize(nil);
End;

Procedure TfEtiqueta.sgPropertiesMouseDown(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);
Var
  LCol, LRow: Integer;
  LPropName: String;
  LNewVal: String;
  LSG: TStringGrid;
Begin
  LSG := ASender AS TStringGrid;
  LSG.MouseToCell(AX, AY, LCol, LRow);
  If (LRow >= 0) AND (LRow < LSG.RowCount) Then
  Begin
    LPropName := LSG.Cells[0, LRow];

    If LCol = 0 Then
    Begin
      If (Pos('Margens', LPropName) > 0) Then
      Begin
        FMarginsExpanded := NOT FMarginsExpanded;
        // Refresh
        If NOT Assigned(FSelectedItem) Then
          ShowLabelProperties;
      End
      Else If (Pos('Bordas', LPropName) > 0) Then
      Begin
        FBordersExpanded := NOT FBordersExpanded;
        // Refresh
        If Assigned(FSelectedItem) Then
          UpdatePropertiesPanel(FSelectedItem);
      End
      Else If (Pos('Célula', LPropName) > 0) Then
      Begin
        FCellPropsExpanded := NOT FCellPropsExpanded;
        // Refresh
        If Assigned(FSelectedItem) Then
          UpdatePropertiesPanel(FSelectedItem);
      End
      Else If (Pos('Grelha', LPropName) > 0) Then
      Begin
        FGridPropsExpanded := NOT FGridPropsExpanded;
        // Refresh
        If Assigned(FSelectedItem) Then
          UpdatePropertiesPanel(FSelectedItem);
      End
      Else If (Pos('Mosaico', LPropName) > 0) Then
      Begin
        FMosaicoExpanded := NOT FMosaicoExpanded;
        // Refresh
        If NOT Assigned(FSelectedItem) Then
          ShowLabelProperties;
      End
      Else If (Pos('Font', LPropName) > 0) Then
      Begin
        FFontExpanded := NOT FFontExpanded;
        If Assigned(FSelectedItem) Then
          UpdatePropertiesPanel(FSelectedItem);
      End
      Else If (Pos('Estilo', LPropName) > 0) Then
      Begin
        FFontStyleExpanded := NOT FFontStyleExpanded;
        If Assigned(FSelectedItem) Then
          UpdatePropertiesPanel(FSelectedItem);
      End
      Else If (Trim(LPropName) = 'Imagem') Then
      Begin
        // Abrir diálogo para escolher imagem ao clicar no nome da propriedade
        If Assigned(FSelectedItem) AND (FSelectedItem IS TLabelImage) Then
        Begin
          If opdImagem.Execute Then
          Begin
            Try
              TLabelImage(FSelectedItem).Picture.LoadFromFile(opdImagem.FileName);
            Except
              TLabelImage(FSelectedItem).Picture.Clear;
            End;
            UpdatePropertiesPanel(FSelectedItem);
            GetActivePaintBox.Invalidate;
          End;
        End;
      End;
    End
    Else If LCol = 1 Then
    Begin
      // Handle Checkbox clicks
      If (LPropName = 'Mostrar Texto') OR (Trim(LPropName) = 'Esquerda') OR (Trim(LPropName) = 'Topo') OR (Trim(LPropName) = 'Direita') OR
        (Trim(LPropName) = 'Baixo') OR (Trim(LPropName) = 'Ajustar') OR (Trim(LPropName) = 'Proporcional') Then
      Begin
        If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
        Begin
          // It's a boolean property
          If SameText(LSG.Cells[1, LRow], 'True') Then
            LNewVal := 'False'
          Else
            LNewVal := 'True';

          LSG.Cells[1, LRow] := LNewVal;
          sgPropertiesSetEditText(LSG, 1, LRow, LNewVal);
        End
        Else If ((LPropName = 'Mostrar Texto') AND (FSelectedItem <> nil) AND (FSelectedItem IS TLabelBarcode)) OR
          (((Trim(LPropName) = 'Ajustar') OR (Trim(LPropName) = 'Proporcional')) AND (FSelectedItem <> nil) AND (FSelectedItem IS TLabelImage)) Then
        Begin
          If SameText(LSG.Cells[1, LRow], 'True') Then
            LNewVal := 'False'
          Else
            LNewVal := 'True';

          LSG.Cells[1, LRow] := LNewVal;
          sgPropertiesSetEditText(LSG, 1, LRow, LNewVal);
        End;
      End;
      // Abrir diálogo para escolher imagem ao clicar no valor da propriedade
      If (Trim(LPropName) = 'Imagem') AND Assigned(FSelectedItem) AND (FSelectedItem IS TLabelImage) Then
      Begin
        If opdImagem.Execute Then
        Begin
          Try
            TLabelImage(FSelectedItem).Picture.LoadFromFile(opdImagem.FileName);
          Except
            TLabelImage(FSelectedItem).Picture.Clear;
          End;
          UpdatePropertiesPanel(FSelectedItem);
          GetActivePaintBox.Invalidate;
        End;
      End;
    End;
  End;
End;

Procedure TfEtiqueta.sgPropertiesDrawCell(ASender: TObject; ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState);
Var
  LPropName, LVal: String;
  LIsBool: Boolean;
  LCheckRect: TRect;
  LState: Integer;
  LSG: TStringGrid;
  LSettings: TTabSettings;
Begin
  If ARow = 0 Then Exit;
  LSG := ASender AS TStringGrid;
  LSettings := GetActiveSettings;

  LPropName := LSG.Cells[0, ARow];
  LVal := LSG.Cells[ACol, ARow];

  // Check if it's a header row
  If (Pos(sMais, LPropName) > 0) OR (Pos(sMenos, LPropName) > 0) Then
  Begin
    LSG.Canvas.Brush.Color := $00404040; // Dark Gray
    LSG.Canvas.Font.Color := clWhite;
    LSG.Canvas.Font.Style := [fsBold];
    LSG.Canvas.FillRect(ARect);

    // Only draw text in the first column, or draw it across if we want to simulate merging
    If ACol = 0 Then
      LSG.Canvas.TextOut(ARect.Left + 2, ARect.Top + 2, LPropName);
    Exit;
  End;

  // Visual indication for locked dimensions
  If SameText(Trim(LPropName), 'Largura (mm)') Then
  Begin
    If (LSettings <> nil) AND (IsPaperSizeContinuous(LSettings.PaperName) OR NOT IsPaperSizeEditable(LSettings.PaperName)) Then
      LSG.Canvas.Brush.Color := clSilver
    Else
      LSG.Canvas.Brush.Color := clWindow;
  End
  Else If SameText(Trim(LPropName), 'Altura (mm)') Then
  Begin
    If (LSettings <> nil) AND (NOT IsPaperSizeContinuous(LSettings.PaperName) AND NOT IsPaperSizeEditable(LSettings.PaperName)) Then
      LSG.Canvas.Brush.Color := clSilver
    Else
      LSG.Canvas.Brush.Color := clWindow;
  End
  Else
    LSG.Canvas.Brush.Color := clWindow;

  LSG.Canvas.FillRect(ARect);

  LIsBool := False;
  If (ACol = 1) Then
  Begin
    If (LPropName = 'Mostrar Texto') OR (Trim(LPropName) = 'Esquerda') OR (Trim(LPropName) = 'Topo') OR (Trim(LPropName) = 'Direita') OR
      (Trim(LPropName) = 'Baixo') OR (Trim(LPropName) = 'Ajustar') OR (Trim(LPropName) = 'Proporcional') Then
    Begin
      If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then LIsBool := True
      Else If (LPropName = 'Mostrar Texto') AND (FSelectedItem <> nil) AND (FSelectedItem IS TLabelBarcode) Then LIsBool := True
      Else If ((Trim(LPropName) = 'Ajustar') OR (Trim(LPropName) = 'Proporcional')) AND (FSelectedItem <> nil) AND (FSelectedItem IS TLabelImage) Then LIsBool := True;
    End;
  End;

  If LIsBool Then
  Begin
    // Draw Checkbox
    LCheckRect := Rect(ARect.Left + 2, ARect.Top + 2, ARect.Left + 18, ARect.Bottom - 2);
    If LCheckRect.Bottom - LCheckRect.Top > 16 Then LCheckRect.Bottom := LCheckRect.Top + 16;
    If LCheckRect.Right - LCheckRect.Left > 16 Then LCheckRect.Right := LCheckRect.Left + 16;

    // Center vertically
    OffsetRect(LCheckRect, 0, (ARect.Bottom - ARect.Top - 16) DIV 2);

    LState := DFCS_BUTTONCHECK;
    If SameText(LVal, 'True') Then LState := LState OR DFCS_CHECKED;

    DrawFrameControl(LSG.Canvas.Handle, LCheckRect, DFC_BUTTON, LState);

    // Draw Text
    LSG.Canvas.TextOut(LCheckRect.Right + 4, ARect.Top + 2, LVal);
  End
  Else
  Begin
    LSG.Canvas.TextOut(ARect.Left + 2, ARect.Top + 2, LVal);
  End;
End;

Procedure TfEtiqueta.UpdateLabelSize;
Var
  LPWidth, LPHeight: Integer;
  LPDPIX, LPDPIY: Integer;
  LSWidth, LSHeight: Integer;
  LTemp: Integer;
  LPB: TPaintBox;
  LSettings: TTabSettings;
Begin
  LSettings := GetActiveSettings;
  If LSettings = nil Then Exit;
  If LSettings.PrinterName = '' Then Exit;

  LPB := GetActivePaintBox;
  If LPB = nil Then Exit;

  Try
    Printer.SetPrinter(LSettings.PrinterName);
    Printer.PaperSize.PaperName := LSettings.PaperName;

    If (LSettings.PaperWidth > 0) AND (LSettings.PaperHeight > 0) Then
    Begin
      LSWidth := LSettings.PaperWidth;
      LSHeight := LSettings.PaperHeight;
      LPDPIX := Screen.PixelsPerInch; // Base for converted values
      LPDPIY := Screen.PixelsPerInch;
    End
    Else
    Begin
      // Get dimensions in printer pixels
      LPWidth := Printer.PaperSize.Width;
      LPHeight := Printer.PaperSize.Height;
      LPDPIX := Printer.XDPI;
      LPDPIY := Printer.YDPI;

      If (LPDPIX = 0) OR (LPDPIY = 0) Then Exit;

      // Convert to Screen Pixels (Screen.PixelsPerInch is usually 96)
      LSWidth := Round(LPWidth / LPDPIX * Screen.PixelsPerInch);
      LSHeight := Round(LPHeight / LPDPIY * Screen.PixelsPerInch);
      
      // Store detected size
      LSettings.PaperWidth := LSWidth;
      LSettings.PaperHeight := LSHeight;
    End;

    // Get hardware margins (unprintable area) only if not loaded from file
    If NOT FDataLoadedOnce Then
    Begin
      With Printer.PaperSize.PaperRect Do
      Begin
        LSettings.Margins.Left := Round((WorkRect.Left - PhysicalRect.Left) / LPDPIX * Screen.PixelsPerInch);
        LSettings.Margins.Top := Round((WorkRect.Top - PhysicalRect.Top) / LPDPIY * Screen.PixelsPerInch);
        LSettings.Margins.Right := Round((PhysicalRect.Right - WorkRect.Right) / LPDPIX * Screen.PixelsPerInch);
        LSettings.Margins.Bottom := Round((PhysicalRect.Bottom - WorkRect.Bottom) / LPDPIY * Screen.PixelsPerInch);
      End;
    End;

    // Apply Orientation
    If LSettings.Orientation = poLandscape Then
    Begin
      If LSHeight > LSWidth Then
      Begin
        LTemp := LSWidth;
        LSWidth := LSHeight;
        LSHeight := LTemp;
      End;
    End
    Else
    Begin
      If LSWidth > LSHeight Then
      Begin
        LTemp := LSWidth;
        LSWidth := LSHeight;
        LSHeight := LTemp;
      End;
    End;

    // Auto-detect multiple labels (e.g. "54mm x 3")
    LTemp := Pos(' X ', UpperCase(LSettings.PaperName));
    If LTemp > 0 Then
    Begin
      LTemp := StrToIntDef(Trim(Copy(LSettings.PaperName, LTemp + 3, 5)), 0);
      If (LTemp > 0) AND (Pos('MM', UpperCase(Copy(LSettings.PaperName, LTemp + 3, 20))) = 0) Then
        LSettings.LayoutCols := LTemp;
    End;

    // Store physical paper size
    LSettings.PaperWidth := LSWidth;
    LSettings.PaperHeight := LSHeight;

    // Expand canvas (TPaintBox) to show objects outside the page
    LPB.Width := LSWidth;
    LPB.Height := LSHeight;

    If Assigned(LabelManager) Then
    Begin
      For LTemp := 0 To LabelManager.Count - 1 Do
      Begin
        If LabelManager[LTemp].Left + LabelManager[LTemp].Width > LPB.Width Then
          LPB.Width := LabelManager[LTemp].Left + LabelManager[LTemp].Width;
        If LabelManager[LTemp].Top + LabelManager[LTemp].Height > LPB.Height Then
          LPB.Height := LabelManager[LTemp].Top + LabelManager[LTemp].Height;
      End;
    End;

    // Add extra space padding
    LPB.Width := LPB.Width + 200;
    LPB.Height := LPB.Height + 200;

  Except
    SysUtils.Beep;
  End;
End;

Procedure TfEtiqueta.sbNewTextClick(ASender: TObject);
Var
  LTxt: TLabelText;
  LGrid: TLabelGrid;
  LP: TPoint;
Begin
  // Check if we are trying to add to an existing grid cell
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    If LGrid.GetSelectionCount = 1 Then
    Begin
      LP := LGrid.GetFirstSelectedCell;
      If (LP.X >= 0) AND (LP.Y >= 0) AND (LGrid.GetCellItem(LP.Y, LP.X) = nil) Then
      Begin
        LTxt := LabelManager.AddText('Texto');
        LabelManager.ExtractItem(LTxt); // Remove from main list

        LTxt.Name := LabelManager.GenerateUniqueName('Texto');
        LTxt.Popup := pmContext;

        LGrid.SetCellItem(LP.Y, LP.X, LTxt);

        // Select it
        LabelManager.DeselectAll;
        LTxt.Selected := True;
        FSelectedItem := LTxt;
        UpdatePropertiesPanel(LTxt);
        RebuildObjectTree;
        SelectInTree(LTxt);
        GetActivePaintBox.Invalidate;
        Exit;
      End;
    End;
  End;

  LTxt := LabelManager.AddText('Texto');
  LTxt.Name := LabelManager.GenerateUniqueName('Texto');
  LTxt.Popup := pmContext;

  // Select it
  LabelManager.DeselectAll;
  LTxt.Selected := True;
  FSelectedItem := LTxt;
  UpdatePropertiesPanel(LTxt);
  RebuildObjectTree;
  SelectInTree(LTxt);
  GetActivePaintBox.Invalidate;
End;

Procedure TfEtiqueta.sbNewImageClick(ASender: TObject);
Var
  LImg: TLabelImage;
  LGrid: TLabelGrid;
  LP: TPoint;
Begin
  // Check if we are trying to add to an existing grid cell
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    If LGrid.GetSelectionCount = 1 Then
    Begin
      LP := LGrid.GetFirstSelectedCell;
      If (LP.X >= 0) AND (LP.Y >= 0) AND (LGrid.GetCellItem(LP.Y, LP.X) = nil) Then
      Begin
        LImg := LabelManager.AddImage;
        LabelManager.ExtractItem(LImg); // Remove from main list

        LImg.Name := LabelManager.GenerateUniqueName('Imagem');
        LImg.Popup := pmContext;

        LGrid.SetCellItem(LP.Y, LP.X, LImg);

        // Select it
        LabelManager.DeselectAll;
        LImg.Selected := True;
        FSelectedItem := LImg;
        UpdatePropertiesPanel(LImg);
        RebuildObjectTree;
        SelectInTree(LImg);
        GetActivePaintBox.Invalidate;
        Exit;
      End;
    End;
  End;

  LImg := LabelManager.AddImage;
  LImg.Name := LabelManager.GenerateUniqueName('Imagem');
  LImg.Popup := pmContext;

  // Select it
  LabelManager.DeselectAll;
  LImg.Selected := True;
  FSelectedItem := LImg;
  UpdatePropertiesPanel(LImg);
  RebuildObjectTree;
  SelectInTree(LImg);
  GetActivePaintBox.Invalidate;
End;

Procedure TfEtiqueta.sbNewBarCodeClick(ASender: TObject);
Var
  LBC: TLabelBarcode;
  LGrid: TLabelGrid;
  LP: TPoint;
Begin
  // Check if we are trying to add to an existing grid cell
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    If LGrid.GetSelectionCount = 1 Then
    Begin
      LP := LGrid.GetFirstSelectedCell;
      If (LP.X >= 0) AND (LP.Y >= 0) AND (LGrid.GetCellItem(LP.Y, LP.X) = nil) Then
      Begin
        LBC := LabelManager.AddBarcode('123456789', bkCode128);
        LabelManager.ExtractItem(LBC); // Remove from main list

        LBC.Name := LabelManager.GenerateUniqueName('Barcode');
        LBC.Popup := pmContext;

        LGrid.SetCellItem(LP.Y, LP.X, LBC);

        // Select it
        LabelManager.DeselectAll;
        LBC.Selected := True;
        FSelectedItem := LBC;
        UpdatePropertiesPanel(LBC);
        RebuildObjectTree;
        SelectInTree(LBC);
        GetActivePaintBox.Invalidate;
        Exit;
      End;
    End;
  End;

  LBC := LabelManager.AddBarcode('123456789', bkCode128);
  LBC.Name := LabelManager.GenerateUniqueName('Barcode');
  LBC.Popup := pmContext;

  // Select it
  LabelManager.DeselectAll;
  LBC.Selected := True;
  FSelectedItem := LBC;
  UpdatePropertiesPanel(LBC);
  RebuildObjectTree;
  SelectInTree(LBC);
  GetActivePaintBox.Invalidate;
End;


Procedure TfEtiqueta.sbNewGridClick(ASender: TObject);
Var
  LGrid, LNewGrid: TLabelGrid;
  LR, LC: Integer;
  LAdded: Boolean;
  LP: TPoint;
Begin
  // Check if we are trying to add to an existing grid
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    LAdded := False;

    // Check selected cell first
    If LGrid.GetSelectionCount = 1 Then
    Begin
      LP := LGrid.GetFirstSelectedCell;
      If (LP.X >= 0) AND (LP.Y >= 0) AND (LGrid.GetCellItem(LP.Y, LP.X) = nil) Then
      Begin
        // Create new grid
        LNewGrid := LabelManager.AddGrid(2, 2);
        // Remove from manager list because it will be owned by the parent grid
        LabelManager.ExtractItem(LNewGrid);

        LNewGrid.Name := 'SubGrid_' + IntToStr(LP.Y) + '_' + IntToStr(LP.X);
        LNewGrid.Popup := pmContext;

        LGrid.SetCellItem(LP.Y, LP.X, LNewGrid);
        LAdded := True;
      End;
    End;

    // Find first empty cell if not added yet
    If NOT LAdded Then
    Begin
      For LR := 0 To LGrid.Rows - 1 Do
      Begin
        For LC := 0 To LGrid.Cols - 1 Do
        Begin
          If LGrid.GetCellItem(LR, LC) = nil Then
          Begin
            // Create new grid
            LNewGrid := LabelManager.AddGrid(2, 2);
            // Remove from manager list because it will be owned by the parent grid
            LabelManager.ExtractItem(LNewGrid);

            LNewGrid.Name := 'SubGrid_' + IntToStr(LR) + '_' + IntToStr(LC);
            LNewGrid.Popup := pmContext;

            LGrid.SetCellItem(LR, LC, LNewGrid);
            LAdded := True;
            Break;
          End;
        End;
        If LAdded Then Break;
      End;
    End;

    If LAdded Then
      GetActivePaintBox.Invalidate
    Else
      ShowMessage('A grelha selecionada não tem células vazias.');

    Exit;
  End;

  LGrid := LabelManager.AddGrid(2, 2);
  LGrid.Name := LabelManager.GenerateUniqueName('Grid');
  LGrid.Left := 10;
  LGrid.Top := 10;
  LGrid.Width := 200;
  LGrid.Height := 100;
  LGrid.Popup := pmContext;

  // Select it
  LabelManager.DeselectAll;
  LGrid.Selected := True;
  FSelectedItem := LGrid;
  UpdatePropertiesPanel(LGrid);
  RebuildObjectTree;
  SelectInTree(LGrid);
  GetActivePaintBox.Invalidate;
End;

Procedure TfEtiqueta.pbDesignPaint(ASender: TObject);
Var
  Li, Lj: Integer;
  Lr, Lc, RX, RY: Integer;
  LPB: TPaintBox;
  LSettings: TTabSettings;

  Procedure DrawHandles(AItem: TLabelItem);
  Var
    Lr, Lc: Integer;
    LGrid: TLabelGrid;
    LSub: TLabelItem;
  Begin
    If AItem.Selected Then
      DrawSelectionHandles(LPB.Canvas, AItem);

    If AItem IS TLabelGrid Then
    Begin
      LGrid := TLabelGrid(AItem);
      For Lr := 0 To LGrid.Rows - 1 Do
        For Lc := 0 To LGrid.Cols - 1 Do
        Begin
          LSub := LGrid.GetCellItem(Lr, Lc);
          If Assigned(LSub) Then
            DrawHandles(LSub);
        End;
    End;
  End;

Begin
  LPB := ASender AS TPaintBox;
  LSettings := GetActiveSettings;
  If LSettings = nil Then Exit;

  // Preencher o fundo do workspace (fora do papel)
  LPB.Canvas.Brush.Color := clGray;
  LPB.Canvas.Brush.Style := bsSolid;
  LPB.Canvas.FillRect(Rect(0, 0, LPB.Width, LPB.Height));

  // Preencher a área da etiqueta (Papel)
  LPB.Canvas.Brush.Color := LPB.Color;
  LPB.Canvas.FillRect(Rect(0, 0, LSettings.PaperWidth, LSettings.PaperHeight));

  // Desenhar a borda da etiqueta
  LPB.Canvas.Pen.Color := clBlack;
  LPB.Canvas.Pen.Style := psDash;
  LPB.Canvas.Brush.Style := bsClear;
  LPB.Canvas.Rectangle(0, 0, LSettings.PaperWidth, LSettings.PaperHeight);

  // Desenhar Margens
  LPB.Canvas.Pen.Color := clSilver;
  LPB.Canvas.Pen.Style := psDot;
  LPB.Canvas.Rectangle(LSettings.Margins.Left, LSettings.Margins.Top, 
                       LSettings.PaperWidth - LSettings.Margins.Right, 
                       LSettings.PaperHeight - LSettings.Margins.Bottom);

  // Desenhar objetos
  If Assigned(LabelManager) Then
  Begin
    LabelManager.DrawAll(LPB.Canvas);

    // Draw selection handles for selected items (recursive)
    For Li := 0 To LabelManager.Count - 1 Do
      DrawHandles(LabelManager[Li]);
  End;

  // Desenhar exemplificação visual do Mosaico (Repetição)
  If (LSettings.LayoutCols > 1) OR (LSettings.LayoutRows > 1) Then
  Begin
    LPB.Canvas.Pen.Color := $00D0D0D0; // Cinza claro
    LPB.Canvas.Pen.Style := psDot;
    LPB.Canvas.Brush.Style := bsClear;

    Li := LSettings.PaperWidth - LSettings.Margins.Left - LSettings.Margins.Right; // WorkWidth
    Lj := LSettings.PaperHeight - LSettings.Margins.Top - LSettings.Margins.Bottom; // WorkHeight

    If (LSettings.LayoutCols > 0) AND (LSettings.LayoutRows > 0) Then
    Begin
      // CellW/H
      // Li is width, Lj is height
      // Usar double para precisão no cálculo
      Try
        For Lr := 0 To LSettings.LayoutRows - 1 Do
          For Lc := 0 To LSettings.LayoutCols - 1 Do
          Begin
            If (Lc = 0) AND (Lr = 0) Then Continue; // Já desenhado pela borda principal/margens

            // Cálculo da posição da repetição
            // Uma etiqueta individual tem o tamanho disponível dividido pelas colunas/linhas menos os gaps
            // Na verdade, no desenho, apenas mostramos as áreas onde as repetições ocorrerão
            // Assumimos que o desenho atual é a primeira etiqueta
            // Posição c, r:
            // X = Margem.Left + c * (LabelW + GapX)
            // Y = Margem.Top + r * (LabelH + GapY)

            // Como não sabemos o LabelW exato (pode ser o que sobrar), calculamos
            // LabelW = (WorkWidth - (Cols - 1) * GapX) / Cols
            // RX := LSettings.Margins.Left + Round(Lc * ((Li - (LSettings.LayoutCols-1)*LSettings.LayoutGapX)/LSettings.LayoutCols + LSettings.LayoutGapX));
            // Simplificado:
            RX := LSettings.Margins.Left + Round(Lc * ( (Li + LSettings.LayoutGapX) / LSettings.LayoutCols ));
            RY := LSettings.Margins.Top + Round(Lr * ( (Lj + LSettings.LayoutGapY) / LSettings.LayoutRows ));
            
            // Desenhar apenas o retângulo da área da etiqueta repetida
            LPB.Canvas.Rectangle(RX, RY, RX + Round((Li - (LSettings.LayoutCols-1)*LSettings.LayoutGapX)/LSettings.LayoutCols),
                                      RY + Round((Lj - (LSettings.LayoutRows-1)*LSettings.LayoutGapY)/LSettings.LayoutRows));
          End;
      Except
      End;
    End;
  End;
End;

Procedure TfEtiqueta.pbDesignMouseDown(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);
Var
  Li: Integer;
  LHandle: TResizeHandle;
  LCol, LRow: Integer;
  LItem, LSubItem: TLabelItem;
  LGrid: TLabelGrid;
  LPB: TPaintBox;
Begin
  LPB := ASender AS TPaintBox;
  If AButton = mbRight Then
  Begin
    LItem := LabelManager.FindItemAt(AX, AY);

    // Drill-down para grelhas se clicarmos numa célula com conteúdo
    If (LItem <> nil) AND (LItem IS TLabelGrid) Then
    Begin
      LGrid := TLabelGrid(LItem);
      If LGrid.GetCellAt(AX, AY, LRow, LCol) Then
      Begin
        LSubItem := LGrid.GetCellItem(LRow, LCol);
        If Assigned(LSubItem) Then
          LItem := LSubItem;
      End;
    End;

    If Assigned(LItem) Then
    Begin
      // Se o item não estiver selecionado, limpamos tudo e selecionamos este
      If NOT LItem.Selected Then
      Begin
        LabelManager.DeselectAll;
        LItem.Selected := True;
      End;

      FSelectedItem := LItem;

      // Se for uma grelha (ou se o item selecionado estiver numa grelha)
      // garantimos que a célula sob o rato está selecionada para o menu de contexto
      If LItem IS TLabelGrid Then
      Begin
        LGrid := TLabelGrid(LItem);
        If LGrid.GetCellAt(AX, AY, LRow, LCol) Then
        Begin
          If NOT LGrid.IsSelected(LRow, LCol) Then
            LGrid.SelectCell(LRow, LCol);
        End;
      End
      Else
      Begin
        // Se for um item dentro de uma grelha, também queremos que a célula esteja selecionada
        // para que as opções de grelha apareçam no menu de contexto
        For Li := 0 To LabelManager.Count - 1 Do
        Begin
          If LabelManager[Li] IS TLabelGrid Then
          Begin
            LGrid := TLabelGrid(LabelManager[Li]);
            If LGrid.GetCellAt(AX, AY, LRow, LCol) Then
            Begin
              If LGrid.GetCellItem(LRow, LCol) = LItem Then
              Begin
                If NOT LGrid.IsSelected(LRow, LCol) Then
                  LGrid.SelectCell(LRow, LCol);
                Break;
              End;
            End;
          End;
        End;
      End;

      UpdatePropertiesPanel(LItem);
      SelectInTree(LItem);
    End
    Else
    Begin
      LabelManager.DeselectAll;
      FSelectedItem := nil;
      ShowLabelProperties;
      SelectInTree(nil);
    End;

    LPB.Invalidate;
    If Assigned(pmContext) Then pmContext.PopUp;
    Exit;
  End;

  If AButton = mbLeft Then
  Begin
    LItem := LabelManager.FindItemAt(AX, AY);

    // 1. Check for grid line resizing (any grid under mouse)
    If (LItem <> nil) AND (LItem IS TLabelGrid) Then
    Begin
      If TLabelGrid(LItem).HitTestGridLine(AX, AY, FResizingIsRow, FResizingLineIndex) Then
      Begin
        // Select it if not selected
        If NOT LItem.Selected Then
        Begin
          LabelManager.DeselectAll;
          LItem.Selected := True;
          FSelectedItem := LItem;
          SelectInTree(LItem);
          UpdatePropertiesPanel(FSelectedItem);
        End;
        FIsResizingGridLine := True;
        FResizingGrid := TLabelGrid(LItem);
        FDragStartPos := Point(AX, AY);
        FIsDragging := False;
        LPB.Invalidate;
        Exit;
      End;
    End;

    // 2. Check if clicking on a handle of selected item
    If Assigned(FSelectedItem) AND FSelectedItem.Selected Then
    Begin
      LHandle := DetectHandle(FSelectedItem, AX, AY);
      If LHandle <> rhNone Then
      Begin
        FIsResizing := True;
        FResizeHandle := LHandle;
        FDragStartPos := Point(AX, AY);
        FResizeStartRect := Rect(FSelectedItem.Left, FSelectedItem.Top, FSelectedItem.Left + FSelectedItem.Width, FSelectedItem.Top + FSelectedItem.Height);
        Exit;
      End;
    End;

    // 3. Drill-down logic: If we clicked a Grid (and not a line)
    If (LItem <> nil) AND (LItem IS TLabelGrid) Then
    Begin
      LGrid := TLabelGrid(LItem);
      If LGrid.GetCellAt(AX, AY, LRow, LCol) Then
      Begin
        LSubItem := LGrid.GetCellItem(LRow, LCol);

        // If there's a sub-item in the cell, select it instead of the grid
        If Assigned(LSubItem) AND NOT (ssCtrl IN AShift) AND NOT (ssShift IN AShift) Then
        Begin
          LabelManager.DeselectAll;
          LSubItem.Selected := True;
          FSelectedItem := LSubItem;
          SelectInTree(LSubItem);
          UpdatePropertiesPanel(LSubItem);
          FDragItem := LSubItem;
          FDragStartPos := Point(AX, AY);
          FIsDragging := True;
          LPB.Invalidate;
          Exit;
        End;

        // If Grid was not selected, select it now
        If NOT LGrid.Selected Then
        Begin
          LabelManager.DeselectAll;
          LGrid.Selected := True;
          FSelectedItem := LGrid;
          SelectInTree(LGrid);
        End;

        // Handle Cell Selection
        If ssShift IN AShift Then
        Begin
          If LGrid.GetSelectionCount > 0 Then
            LGrid.SelectRange(LGrid.GetFirstSelectedCell.Y, LGrid.GetFirstSelectedCell.X, LRow, LCol, True)
          Else
            LGrid.SelectCell(LRow, LCol);
        End
        Else If ssCtrl IN AShift Then
        Begin
          LGrid.ToggleCellSelection(LRow, LCol);
        End
        Else
        Begin
          LGrid.SelectCell(LRow, LCol);
        End;

        UpdatePropertiesPanel(FSelectedItem);
        SelectInTree(FSelectedItem);
        LPB.Invalidate;

        // If we are selecting a cell, we might not want to drag the whole grid immediately
        // unless we click and move. Let's set up for dragging but allow cell selection to take precedence.
        FDragItem := LGrid;
        FDragStartPos := Point(AX, AY);
        FIsDragging := True;
        Exit; // Important: we handled the click
      End;
    End;

    // Handle Selection Logic
    If Assigned(LItem) Then
    Begin
      If ssCtrl IN AShift Then
      Begin
        LItem.Selected := NOT LItem.Selected;
        If LItem.Selected Then
          FSelectedItem := LItem
        Else If FSelectedItem = LItem Then
          FSelectedItem := nil;
      End
      Else
      Begin
        // Exclusive selection
        LabelManager.DeselectAll;
        LItem.Selected := True;
        FSelectedItem := LItem;
      End;

      If Assigned(LItem) AND LItem.Selected Then
      Begin
        FDragItem := LItem;
        FDragStartPos := Point(AX, AY);
        FIsDragging := True;
        UpdatePropertiesPanel(FSelectedItem);
        SelectInTree(LItem);

        // Give focus to container so keyboard shortcuts work
        If GetActiveScrollBox.CanFocus Then
          GetActiveScrollBox.SetFocus;
      End
      Else
      Begin
        // Item was deselected via Ctrl+Click
        UpdatePropertiesPanel(FSelectedItem);
        SelectInTree(FSelectedItem);
      End;
    End
    Else
    Begin
      // Clicked on empty space
      If NOT (ssCtrl IN AShift) Then
      Begin
        LabelManager.DeselectAll;
        FSelectedItem := nil;
      End;
      ShowLabelProperties;
      SelectInTree(nil);
    End;

    LPB.Invalidate;
  End;
End;

Procedure TfEtiqueta.pbDesignMouseMove(ASender: TObject; AShift: TShiftState; AX, AY: Integer);
Var
  LDX, LDY: Integer;
  LNewLeft, LNewTop, LNewWidth, LNewHeight: Integer;
  LItem: TLabelItem;
  LHandle: TResizeHandle;
  LPB: TPaintBox;
Begin
  LPB := ASender AS TPaintBox;
  If FIsResizing AND Assigned(FSelectedItem) Then
  Begin
    LDX := AX - FDragStartPos.X;
    LDY := AY - FDragStartPos.Y;

    LNewLeft := FResizeStartRect.Left;
    LNewTop := FResizeStartRect.Top;
    LNewWidth := FResizeStartRect.Right - FResizeStartRect.Left;
    LNewHeight := FResizeStartRect.Bottom - FResizeStartRect.Top;

    Case FResizeHandle Of
      rhTopLeft:
      Begin
        LNewLeft := FResizeStartRect.Left + LDX;
        LNewTop := FResizeStartRect.Top + LDY;
        LNewWidth := FResizeStartRect.Right - LNewLeft;
        LNewHeight := FResizeStartRect.Bottom - LNewTop;
      End;
      rhTop:
      Begin
        LNewTop := FResizeStartRect.Top + LDY;
        LNewHeight := FResizeStartRect.Bottom - LNewTop;
      End;
      rhTopRight:
      Begin
        LNewTop := FResizeStartRect.Top + LDY;
        LNewWidth := FResizeStartRect.Left + LNewWidth + LDX - FResizeStartRect.Left;
        LNewHeight := FResizeStartRect.Bottom - LNewTop;
      End;
      rhRight:
      Begin
        LNewWidth := LNewWidth + LDX;
      End;
      rhBottomRight:
      Begin
        LNewWidth := LNewWidth + LDX;
        LNewHeight := LNewHeight + LDY;
      End;
      rhBottom:
      Begin
        LNewHeight := LNewHeight + LDY;
      End;
      rhBottomLeft:
      Begin
        LNewLeft := FResizeStartRect.Left + LDX;
        LNewWidth := FResizeStartRect.Right - LNewLeft;
        LNewHeight := LNewHeight + LDY;
      End;
      rhLeft:
      Begin
        LNewLeft := FResizeStartRect.Left + LDX;
        LNewWidth := FResizeStartRect.Right - LNewLeft;
      End;
    End;

    // Apply minimum size
    If LNewWidth < 10 Then LNewWidth := 10;
    If LNewHeight < 10 Then LNewHeight := 10;

    FSelectedItem.Left := LNewLeft;
    FSelectedItem.Top := LNewTop;
    FSelectedItem.Width := LNewWidth;
    FSelectedItem.Height := LNewHeight;

    // Update properties panel in real-time
    If Assigned(FSelectedItem) Then
      UpdatePropertiesPanel(FSelectedItem);

    LPB.Invalidate;
  End
  Else If FIsDragging AND Assigned(FDragItem) Then
  Begin
    LDX := AX - FDragStartPos.X;
    LDY := AY - FDragStartPos.Y;

    FDragItem.Left := FDragItem.Left + LDX;
    FDragItem.Top := FDragItem.Top + LDY;

    FDragStartPos := Point(AX, AY);

    // Update properties panel in real-time
    If Assigned(FSelectedItem) Then
      UpdatePropertiesPanel(FSelectedItem);

    LPB.Invalidate;
  End
  Else If FIsResizingGridLine AND Assigned(FResizingGrid) Then
  Begin
    LDX := AX - FDragStartPos.X;
    LDY := AY - FDragStartPos.Y;

    If FResizingIsRow Then
    Begin
      // Resize Row
      If Abs(LDY) > 0 Then
      Begin
        LNewHeight := FResizingGrid.GetRowHeight(FResizingLineIndex) + LDY;
        // Limit: both rows must be at least 10px
        If (LNewHeight >= 10) AND (FResizingGrid.GetRowHeight(FResizingLineIndex + 1) - LDY >= 10) Then
        Begin
          FResizingGrid.SetRowHeight(FResizingLineIndex, LNewHeight);
          FResizingGrid.SetRowHeight(FResizingLineIndex + 1, FResizingGrid.GetRowHeight(FResizingLineIndex + 1) - LDY);
          FDragStartPos.Y := AY;
        End;
      End;
    End
    Else
    Begin
      // Resize Column
      If Abs(LDX) > 0 Then
      Begin
        LNewWidth := FResizingGrid.GetColWidth(FResizingLineIndex) + LDX;
        // Limit: both columns must be at least 10px
        If (LNewWidth >= 10) AND (FResizingGrid.GetColWidth(FResizingLineIndex + 1) - LDX >= 10) Then
        Begin
          FResizingGrid.SetColWidth(FResizingLineIndex, LNewWidth);
          FResizingGrid.SetColWidth(FResizingLineIndex + 1, FResizingGrid.GetColWidth(FResizingLineIndex + 1) - LDX);
          FDragStartPos.X := AX;
        End;
      End;
    End;
    LPB.Invalidate;
  End
  Else
  Begin
    // Update Cursor
    LItem := LabelManager.FindItemAt(AX, AY);
    If (LItem <> nil) AND (LItem IS TLabelGrid) Then
    Begin
      If TLabelGrid(LItem).HitTestGridLine(AX, AY, FResizingIsRow, FResizingLineIndex) Then
      Begin
        If FResizingIsRow Then LPB.Cursor := crVSplit
        Else
          LPB.Cursor := crHSplit;
      End
      Else
      Begin
        LHandle := DetectHandle(FSelectedItem, AX, AY);
        If LHandle <> rhNone Then
        Begin
          Case LHandle Of
            rhTopLeft, rhBottomRight: LPB.Cursor := crSizeNWSE;
            rhTopRight, rhBottomLeft: LPB.Cursor := crSizeNESW;
            rhTop, rhBottom: LPB.Cursor := crSizeNS;
            rhLeft, rhRight: LPB.Cursor := crSizeWE;
          End;
        End
        Else
          LPB.Cursor := crDefault;
      End;
    End
    Else
    Begin
      LHandle := DetectHandle(FSelectedItem, AX, AY);
      If LHandle <> rhNone Then
      Begin
        Case LHandle Of
          rhTopLeft, rhBottomRight: LPB.Cursor := crSizeNWSE;
          rhTopRight, rhBottomLeft: LPB.Cursor := crSizeNESW;
          rhTop, rhBottom: LPB.Cursor := crSizeNS;
          rhLeft, rhRight: LPB.Cursor := crSizeWE;
        End;
      End
      Else
        LPB.Cursor := crDefault;
    End;
  End;
End;

Procedure TfEtiqueta.pbDesignMouseUp(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);
Begin
  FIsDragging := False;
  FIsResizing := False;
  FIsResizingGridLine := False;
  FDragItem := nil;
  FResizingGrid := nil;
  FResizeHandle := rhNone;
  UpdateLabelSize;
End;

Procedure TfEtiqueta.sbAlignObjectClick(ASender: TObject);
Var
  LSelectedItems: TList;
  Li: Integer;
  LItem: TLabelItem;
  LRefRect: TRect;
  LAlignToPage: Boolean;
  LFirstItem: TLabelItem;
  LSettings: TTabSettings;
Begin
  LSettings := GetActiveSettings;
  If LSettings = nil Then Exit;

  LSelectedItems := LabelManager.GetSelectedItems;
  Try
    If LSelectedItems.Count = 0 Then Exit;

    LAlignToPage := (cbAlignTo.ItemIndex = 0) OR (LSelectedItems.Count = 1);

    If LAlignToPage Then
    Begin
      // Align to Page Margins
      LRefRect := Rect(LSettings.Margins.Left, LSettings.Margins.Top, LSettings.PaperWidth - LSettings.Margins.Right, LSettings.PaperHeight - LSettings.Margins.Bottom);
    End
    Else
    Begin
      // Align to 1st Selected Object
      LFirstItem := TLabelItem(LSelectedItems[0]);
      LRefRect := Rect(LFirstItem.Left, LFirstItem.Top, LFirstItem.Left + LFirstItem.Width, LFirstItem.Top + LFirstItem.Height);
    End;

    For Li := 0 To LSelectedItems.Count - 1 Do
    Begin
      LItem := TLabelItem(LSelectedItems[Li]);
      // If aligning to first item, skip the first item itself
      If (NOT LAlignToPage) AND (Li = 0) Then Continue;

      If ASender = sbAlignObjectLeft Then
        LItem.Left := LRefRect.Left
      Else If ASender = sbAlignObjectRight Then
        LItem.Left := LRefRect.Right - LItem.Width
      Else If ASender = sbAlignObjectTop Then
        LItem.Top := LRefRect.Top
      Else If ASender = sbAlignObjectBottom Then
        LItem.Top := LRefRect.Bottom - LItem.Height
      Else If ASender = sbAlignObjectCenterHorizontal Then
        LItem.Top := LRefRect.Top + (LRefRect.Height - LItem.Height) DIV 2
      Else If ASender = sbAlignObjectCenterVertical Then
        LItem.Left := LRefRect.Left + (LRefRect.Width - LItem.Width) DIV 2;
    End;

    GetActivePaintBox.Invalidate;
    If Assigned(FSelectedItem) Then
      UpdatePropertiesPanel(FSelectedItem);
  Finally
    LSelectedItems.Free;
  End;
End;

Procedure TfEtiqueta.sbExpandMarginsClick(ASender: TObject);
Var
  LSelectedItems: TList;
  Li, Lr, Lc, Lj: Integer;
  LItem: TLabelItem;
  LGrid: TLabelGrid;
  LTargetRect: TRect;
  LFound: Boolean;
  LMerge: TMergedCell;
  CellW, CellH: Double;
  LSettings: TTabSettings;
Begin
  LSettings := GetActiveSettings;
  If LSettings = nil Then Exit;

  // Initialize LMerge to avoid hints
  LMerge.Row := -1;
  LMerge.Col := -1;
  LMerge.RowSpan := 0;
  LMerge.ColSpan := 0;

  LSelectedItems := LabelManager.GetSelectedItems;
  Try
    If LSelectedItems.Count = 0 Then Exit;

    For Li := 0 To LSelectedItems.Count - 1 Do
    Begin
      LItem := TLabelItem(LSelectedItems[Li]);
      LFound := False;

      // 1. Verificar se o objeto está dentro de uma grelha
      For Lj := 0 To LabelManager.Count - 1 Do
      Begin
        If LabelManager[Lj] IS TLabelGrid Then
        Begin
          LGrid := TLabelGrid(LabelManager[Lj]);
          For Lr := 0 To LGrid.Rows - 1 Do
          Begin
            For Lc := 0 To LGrid.Cols - 1 Do
            Begin
              If LGrid.GetCellItem(Lr, Lc) = LItem Then
              Begin
                LFound := True;
                CellW := LGrid.Width / LGrid.Cols;
                CellH := LGrid.Height / LGrid.Rows;

                If LGrid.IsCellMerged(Lr, Lc, LMerge) Then
                Begin
                  LTargetRect.Left := LGrid.Left + Round(LMerge.Col * CellW);
                  LTargetRect.Top := LGrid.Top + Round(LMerge.Row * CellH);
                  LTargetRect.Right := LGrid.Left + Round((LMerge.Col + LMerge.ColSpan) * CellW);
                  LTargetRect.Bottom := LGrid.Top + Round((LMerge.Row + LMerge.RowSpan) * CellH);
                End
                Else
                Begin
                  LTargetRect.Left := LGrid.Left + Round(Lc * CellW);
                  LTargetRect.Top := LGrid.Top + Round(Lr * CellH);
                  LTargetRect.Right := LGrid.Left + Round((Lc + 1) * CellW);
                  LTargetRect.Bottom := LGrid.Top + Round((Lr + 1) * CellH);
                End;
                Break;
              End;
            End;
            If LFound Then Break;
          End;
        End;
        If LFound Then Break;
      End;

      If LFound Then
      Begin
        LItem.Left := LTargetRect.Left;
        LItem.Top := LTargetRect.Top;
        LItem.Width := LTargetRect.Right - LTargetRect.Left;
        LItem.Height := LTargetRect.Bottom - LTargetRect.Top;
      End
      Else
      Begin
        // 2. Expandir para as margens da página
        LItem.Left := LSettings.Margins.Left;
        LItem.Top := LSettings.Margins.Top;
        LItem.Width := LSettings.PaperWidth - LSettings.Margins.Left - LSettings.Margins.Right;
        LItem.Height := LSettings.PaperHeight - LSettings.Margins.Top - LSettings.Margins.Bottom;
      End;
    End;

    GetActivePaintBox.Invalidate;
    If Assigned(FSelectedItem) Then
      UpdatePropertiesPanel(FSelectedItem);
  Finally
    LSelectedItems.Free;
  End;
End;

Procedure TfEtiqueta.sbEqualizeWidthClick(ASender: TObject);
Var
  LSelectedItems: TList;
  LItem, LFirstItem: TLabelItem;
  i, r, c: Integer;
  LGrid: TLabelGrid;
  LTotalWidth, LAvgWidth, LCurrentTotal: Integer;
  LSelectedCols: TList;
  LColIndex: Integer;
Begin
  LSelectedItems := LabelManager.GetSelectedItems;
  Try
    // Case 1: Multiple Objects Selected
    If LSelectedItems.Count > 1 Then
    Begin
      LFirstItem := TLabelItem(LSelectedItems[0]);
      For i := 1 To LSelectedItems.Count - 1 Do
      Begin
        LItem := TLabelItem(LSelectedItems[i]);
        LItem.Width := LFirstItem.Width;
      End;
      GetActivePaintBox.Invalidate;
      Exit;
    End;

    // Case 2: Single Object Selected (Check for Grid Cells)
    If LSelectedItems.Count = 1 Then
    Begin
      LItem := TLabelItem(LSelectedItems[0]);
      If LItem IS TLabelGrid Then
      Begin
        LGrid := TLabelGrid(LItem);
        If LGrid.GetSelectionCount > 1 Then
        Begin
          LSelectedCols := TList.Create;
          Try
            // Identify unique selected columns
            For c := 0 To LGrid.Cols - 1 Do
            Begin
              For r := 0 To LGrid.Rows - 1 Do
              Begin
                If LGrid.IsSelected(r, c) Then
                Begin
                  LSelectedCols.Add(Pointer(PtrInt(c)));
                  Break; // Column is selected
                End;
              End;
            End;

            If LSelectedCols.Count > 1 Then
            Begin
              // Calculate total width of selected columns
              LTotalWidth := 0;
              For i := 0 To LSelectedCols.Count - 1 Do
              Begin
                LColIndex := PtrInt(LSelectedCols[i]);
                LTotalWidth := LTotalWidth + LGrid.GetColWidth(LColIndex);
              End;

              // Distribute evenly
              LAvgWidth := LTotalWidth DIV LSelectedCols.Count;
              LCurrentTotal := 0;

              For i := 0 To LSelectedCols.Count - 1 Do
              Begin
                LColIndex := PtrInt(LSelectedCols[i]);
                If i = LSelectedCols.Count - 1 Then
                  // Last column takes the remainder to ensure exact total match
                  LGrid.SetColWidth(LColIndex, LTotalWidth - LCurrentTotal)
                Else
                Begin
                  LGrid.SetColWidth(LColIndex, LAvgWidth);
                  LCurrentTotal := LCurrentTotal + LAvgWidth;
                End;
              End;
              GetActivePaintBox.Invalidate;
            End;
          Finally
            LSelectedCols.Free;
          End;
        End;
      End;
    End;
  Finally
    LSelectedItems.Free;
  End;
End;

Procedure TfEtiqueta.sbEqualizeHeightClick(ASender: TObject);
Var
  LSelectedItems: TList;
  LItem, LFirstItem: TLabelItem;
  i, r, c: Integer;
  LGrid: TLabelGrid;
  LTotalHeight, LAvgHeight, LCurrentTotal: Integer;
  LSelectedRows: TList;
  LRowIndex: Integer;
Begin
  LSelectedItems := LabelManager.GetSelectedItems;
  Try
    // Case 1: Multiple Objects Selected
    If LSelectedItems.Count > 1 Then
    Begin
      LFirstItem := TLabelItem(LSelectedItems[0]);
      For i := 1 To LSelectedItems.Count - 1 Do
      Begin
        LItem := TLabelItem(LSelectedItems[i]);
        LItem.Height := LFirstItem.Height;
      End;
      GetActivePaintBox.Invalidate;
      Exit;
    End;

    // Case 2: Single Object Selected (Check for Grid Cells)
    If LSelectedItems.Count = 1 Then
    Begin
      LItem := TLabelItem(LSelectedItems[0]);
      If LItem IS TLabelGrid Then
      Begin
        LGrid := TLabelGrid(LItem);
        If LGrid.GetSelectionCount > 1 Then
        Begin
          LSelectedRows := TList.Create;
          Try
            // Identify unique selected rows
            For r := 0 To LGrid.Rows - 1 Do
            Begin
              For c := 0 To LGrid.Cols - 1 Do
              Begin
                If LGrid.IsSelected(r, c) Then
                Begin
                  LSelectedRows.Add(Pointer(PtrInt(r)));
                  Break; // Row is selected
                End;
              End;
            End;

            If LSelectedRows.Count > 1 Then
            Begin
              // Calculate total height of selected rows
              LTotalHeight := 0;
              For i := 0 To LSelectedRows.Count - 1 Do
              Begin
                LRowIndex := PtrInt(LSelectedRows[i]);
                LTotalHeight := LTotalHeight + LGrid.GetRowHeight(LRowIndex);
              End;

              // Distribute evenly
              LAvgHeight := LTotalHeight DIV LSelectedRows.Count;
              LCurrentTotal := 0;

              For i := 0 To LSelectedRows.Count - 1 Do
              Begin
                LRowIndex := PtrInt(LSelectedRows[i]);
                If i = LSelectedRows.Count - 1 Then
                  // Last row takes the remainder
                  LGrid.SetRowHeight(LRowIndex, LTotalHeight - LCurrentTotal)
                Else
                Begin
                  LGrid.SetRowHeight(LRowIndex, LAvgHeight);
                  LCurrentTotal := LCurrentTotal + LAvgHeight;
                End;
              End;
              GetActivePaintBox.Invalidate;
            End;
          Finally
            LSelectedRows.Free;
          End;
        End;
      End;
    End;
  Finally
    LSelectedItems.Free;
  End;
End;

Procedure TfEtiqueta.sbContainerMouseDown(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);
Var
  Li: Integer;
Begin
  If AButton = mbLeft Then
  Begin
    // Clicked on empty space (outside page)
    If NOT (ssCtrl IN AShift) Then
    Begin
      For Li := 0 To LabelManager.Count - 1 Do
      Begin
        LabelManager[Li].Selected := False;
        // Clear cell selections in grids
        If LabelManager[Li] IS TLabelGrid Then
          TLabelGrid(LabelManager[Li]).ClearSelection;
      End;
      FSelectedItem := nil;
    End;
    ShowLabelProperties;
    SelectInTree(nil);
    GetActivePaintBox.Invalidate;
  End;
End;

Procedure TfEtiqueta.pObjetPlusResize(ASender: TObject);
Var
  LSG: TStringGrid;
Begin
  LSG := GetActivePropertiesGrid;
  If LSG = nil Then Exit;

  LSG.AutoSizeColumn(0);
  LSG.ColWidths[1] := LSG.Width - LSG.ColWidths[0] - 16;
End;

Procedure TfEtiqueta.sbAutoSizeColsClick(ASender: TObject);
Begin
  sgData.AutoSizeColumns;
  sgData.ColWidths[0] := iCheckColWidth;
  sgData.AjustarLarguraAoLimite;
End;

Procedure TfEtiqueta.sbExternalDataClick(ASender: TObject);
Begin
  If fDadosExternos.GetDados(EtiquetaFile.Dados) = mrOk Then
  Begin
    FDataLoadedOnce := False; // Resetar para permitir novo AutoSize
    LoadExternalData;
  End;
End;

Procedure TfEtiqueta.sbRefreshDataClick(ASender: TObject);
Begin
  LoadExternalData;
End;

Procedure TfEtiqueta.miDataCheckAllClick(ASender: TObject);
Var
  Li: Integer;
Begin
  sgData.BeginUpdate;
  Try
    For Li := 2 To sgData.RowCount - 1 Do
      sgData.Cells[0, Li] := '1';
  Finally
    sgData.EndUpdate;
  End;
End;

Procedure TfEtiqueta.miDataUncheckAllClick(ASender: TObject);
Var
  Li: Integer;
Begin
  sgData.BeginUpdate;
  Try
    For Li := 2 To sgData.RowCount - 1 Do
      sgData.Cells[0, Li] := '0';
  Finally
    sgData.EndUpdate;
  End;
End;

Procedure TfEtiqueta.tmrFilterTimer(ASender: TObject);
Begin
  tmrFilter.Enabled := False;
  LoadExternalData;
End;

Procedure TfEtiqueta.sbSaveFileClick(ASender: TObject);
Begin
  If EtiquetaFile.FileName = '' Then
  Begin
    If fMain.sdMain.Execute Then
      EtiquetaFile.FileName := WideString(fMain.sdMain.FileName)
    Else
      Exit;
  End;

  UpdateEtiquetaRecord;
  EtiquetaFile.SaveToFile(EtiquetaFile.FileName);
  Caption := ExtractFileName(EtiquetaFile.FileName);
End;

Procedure TfEtiqueta.UpdateEtiquetaRecord;
Var
  i, j, LCount: Integer;
  LManager: TLabelManager;
  LSettings: TTabSettings;
Begin
  // Count valid tabs
  LCount := 0;
  For i := 0 To pcEtiquetas.PageCount - 1 Do
    If (pcEtiquetas.Pages[i].Tag <> 0) Then
      Inc(LCount);

  EtiquetaFile.Modelos := nil;
  SetLength(EtiquetaFile.Modelos, LCount);

  LCount := 0;
  For i := 0 To pcEtiquetas.PageCount - 1 Do
  Begin
    If (pcEtiquetas.Pages[i].Tag = 0) OR (pcEtiquetas.Pages[i] = tsEtiqueta0) Then Continue;

    LSettings := TTabSettings(pcEtiquetas.Pages[i].Tag);
    LManager := LSettings.LabelManager;
    
    EtiquetaFile.Modelos[LCount].Clear;
    EtiquetaFile.Modelos[LCount].Titulo := pcEtiquetas.Pages[i].Caption;
    EtiquetaFile.Modelos[LCount].Impressora := LSettings.PrinterName;
    EtiquetaFile.Modelos[LCount].Papel := LSettings.PaperName;
    EtiquetaFile.Modelos[LCount].Tamanho.Largura := PxToMm(LSettings.PaperWidth);
    EtiquetaFile.Modelos[LCount].Tamanho.Altura := PxToMm(LSettings.PaperHeight);

    EtiquetaFile.Modelos[LCount].Margens.Esquerda := PxToMm(LSettings.Margins.Left);
    EtiquetaFile.Modelos[LCount].Margens.Cima := PxToMm(LSettings.Margins.Top);
    EtiquetaFile.Modelos[LCount].Margens.Direita := PxToMm(LSettings.Margins.Right);
    EtiquetaFile.Modelos[LCount].Margens.Baixo := PxToMm(LSettings.Margins.Bottom);
    EtiquetaFile.Modelos[LCount].Orientacao := Ord(LSettings.Orientation);

    EtiquetaFile.Modelos[LCount].Layout.Cols := LSettings.LayoutCols;
    EtiquetaFile.Modelos[LCount].Layout.Rows := LSettings.LayoutRows;
    EtiquetaFile.Modelos[LCount].Layout.GapX := PxToMm(LSettings.LayoutGapX);
    EtiquetaFile.Modelos[LCount].Layout.GapY := PxToMm(LSettings.LayoutGapY);

    // Objects
    SetLength(EtiquetaFile.Modelos[LCount].Objetos, LManager.Count);
    For j := 0 To LManager.Count - 1 Do
    Begin
      EtiquetaFile.Modelos[LCount].Objetos[j].Objeto := LManager[j];
    End;
    
    Inc(LCount);
  End;
End;

Procedure TfEtiqueta.LoadFromEtiquetaRecord;
Var
  i, j: Integer;
  LManager: TLabelManager;
  LSettings: TTabSettings;
Begin
  // Clear existing tabs (except tsMais and tsEtiqueta0)
  For i := pcEtiquetas.PageCount - 1 DownTo 0 Do
  Begin
    If (pcEtiquetas.Pages[i] <> tsMais) AND (pcEtiquetas.Pages[i] <> tsEtiqueta0) Then
    Begin
      If pcEtiquetas.Pages[i].Tag <> 0 Then
        TTabSettings(pcEtiquetas.Pages[i].Tag).Free;
      pcEtiquetas.Pages[i].Free;
    End;
  End;

  If Length(EtiquetaFile.Modelos) = 0 Then
  Begin
    AddNewTab('Etiqueta 0');
    Exit;
  End;

  For i := 0 To High(EtiquetaFile.Modelos) Do
  Begin
    // Create new tabs for all models
    AddNewTab(EtiquetaFile.Modelos[i].Titulo);
    LSettings := GetActiveSettings;

    LManager := LSettings.LabelManager;

    LSettings.PrinterName := EtiquetaFile.Modelos[i].Impressora;
    LSettings.PaperName := EtiquetaFile.Modelos[i].Papel;
    LSettings.PaperWidth := MmToPx(EtiquetaFile.Modelos[i].Tamanho.Largura);
    LSettings.PaperHeight := MmToPx(EtiquetaFile.Modelos[i].Tamanho.Altura);
    LSettings.Margins.Left := MmToPx(EtiquetaFile.Modelos[i].Margens.Esquerda);
    LSettings.Margins.Top := MmToPx(EtiquetaFile.Modelos[i].Margens.Cima);
    LSettings.Margins.Right := MmToPx(EtiquetaFile.Modelos[i].Margens.Direita);
    LSettings.Margins.Bottom := MmToPx(EtiquetaFile.Modelos[i].Margens.Baixo);

    If EtiquetaFile.Modelos[i].Orientacao = 1 Then
      LSettings.Orientation := poLandscape
    Else
      LSettings.Orientation := poPortrait;

    LSettings.LayoutCols := EtiquetaFile.Modelos[i].Layout.Cols;
    LSettings.LayoutRows := EtiquetaFile.Modelos[i].Layout.Rows;
    LSettings.LayoutGapX := MmToPx(EtiquetaFile.Modelos[i].Layout.GapX);
    LSettings.LayoutGapY := MmToPx(EtiquetaFile.Modelos[i].Layout.GapY);

    // Objects
    For j := 0 To High(EtiquetaFile.Modelos[i].Objetos) Do
    Begin
      LManager.AddItem(EtiquetaFile.Modelos[i].Objetos[j].Objeto);
    End;
  End;

  pcEtiquetas.ActivePageIndex := 0;
  UpdateLabelSize;
  RebuildObjectTree;
  ShowLabelProperties;
  LoadExternalData;
End;

Procedure TfEtiqueta.pcEtiquetasChange(ASender: TObject);
Var
  LPB: TPaintBox;
Begin
  If FIsAddingTab Then Exit;

  If pcEtiquetas.ActivePage = tsMais Then
  Begin
    FIsAddingTab := True;
    Try
      AddNewTab('');
    Finally
      FIsAddingTab := False;
    End;
  End
  Else
  Begin
    FSelectedItem := nil;
    ShowLabelProperties;
    RebuildObjectTree;
    LPB := GetActivePaintBox;
    If LPB <> nil Then LPB.Invalidate;
  End;
End;

Procedure TfEtiqueta.sgDataSelectCell(ASender: TObject; ACol, ARow: Integer; Var ACanSelect: Boolean);
Begin
  ACanSelect := True;
  UpdateObjectsFromData(ARow);
  If ARow = 1 Then FLastFilterValue := sgData.Cells[ACol, ARow];
  If ARow = 1 Then
  Begin
    If ACol = 0 Then // Coluna de checkbox de filtro não é editável por texto
    Begin
      If goEditing IN sgData.Options Then
        sgData.Options := sgData.Options - [goEditing];
    End
    Else
    Begin
      If NOT (goEditing IN sgData.Options) Then
        sgData.Options := sgData.Options + [goEditing];
    End;
  End
  Else
  Begin
    If goEditing IN sgData.Options Then
      sgData.Options := sgData.Options - [goEditing];
  End;
End;

Procedure TfEtiqueta.sgDataMouseDown(ASender: TObject; AButton: TMouseButton; AShift: TShiftState; AX, AY: Integer);
Var
  LCol, LRow: Integer;
Begin
  sgData.MouseToCell(AX, AY, LCol, LRow);
  If (LCol = 0) AND (LRow >= 1) Then
  Begin
    If sgData.Cells[LCol, LRow] = '1' Then
      sgData.Cells[LCol, LRow] := '0'
    Else
      sgData.Cells[LCol, LRow] := '1';

    If LRow = 1 Then
    Begin
      tmrFilter.Enabled := False;
      tmrFilter.Enabled := True;
    End;
  End;
End;

Procedure TfEtiqueta.sgDataSetEditText(ASender: TObject; ACol, ARow: Integer; Const AValue: String);
Begin
  If ARow = 1 Then
  Begin
    If AValue <> FLastFilterValue Then
    Begin
      FLastFilterValue := AValue;
      tmrFilter.Enabled := False;
      tmrFilter.Enabled := True; // Reiniciar timer (debounce)
    End;
  End;
End;

Procedure TfEtiqueta.sgDataDrawCell(ASender: TObject; ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState);
Var
  LFlags: Integer;
Begin
  If (ACol = 0) AND (ARow > 0) Then
  Begin
    // Desenhar Checkbox na Coluna 0
    sgData.Canvas.FillRect(ARect);
    LFlags := DFCS_BUTTONCHECK OR DFCS_ADJUSTRECT;
    If sgData.Cells[ACol, ARow] = '1' Then
      LFlags := LFlags OR DFCS_CHECKED;

    DrawFrameControl(sgData.Canvas.Handle, ARect, DFC_BUTTON, LFlags);
  End
  Else If ARow = 1 Then
  Begin
    // Desenhar apenas a moldura na linha de filtros para parecer uma caixa de edição
    sgData.Canvas.Brush.Style := bsClear;
    sgData.Canvas.Pen.Color := clSilver;
    sgData.Canvas.Rectangle(ARect);
    sgData.Canvas.Brush.Style := bsSolid;
  End;
End;

Procedure TfEtiqueta.sgDataPrepareCanvas(ASender: TObject; ACol, ARow: Integer; AState: TGridDrawState);
Begin
  // Forçar cores na linha de filtros para que o texto nunca fique "invisível"
  If ARow = 1 Then
  Begin
    sgData.Canvas.Brush.Color := clWhite;
    sgData.Canvas.Font.Color := clWindowText;
    sgData.Canvas.Font.Style := [];

    // Se a célula estiver selecionada ou focada, mantemos o fundo branco/claro
    If (gdSelected IN AState) OR (gdFocused IN AState) Then
      sgData.Canvas.Brush.Color := $00F5F5F5;
  End;
End;

Procedure TfEtiqueta.LoadExternalData;
Var
  LFilters: TStringList;
  Lc, Li: Integer;
  LCheckedKeys: TStringList;
  LRowKey: String;
  LSelStart, LSelLen: Integer;
  LEditor: TCustomEdit;
Begin
  StopDataLoader;

  pnlData.Visible := EtiquetaFile.Dados.TipoDadosExternos <> tdeNenhum;
  If pnlData.Visible Then
  Begin
    LFilters := TStringList.Create;
    LCheckedKeys := TStringList.Create;
    Try
      // Recolher chaves das linhas marcadas para as preservar após o reload
      For Li := 2 To sgData.RowCount - 1 Do
      Begin
        If sgData.Cells[0, Li] = '1' Then
        Begin
          LRowKey := '';
          For Lc := 1 To sgData.ColCount - 1 Do
            LRowKey := LRowKey + sgData.Cells[Lc, Li] + '|';
          LCheckedKeys.Add(LRowKey);
        End;
      End;

      // Recolher filtros da Linha 1 (Filtros por coluna)
      For Lc := 1 To sgData.ColCount - 1 Do
      Begin
        If (Trim(sgData.Cells[Lc, 0]) <> '') AND (Trim(sgData.Cells[Lc, 1]) <> '') Then
          LFilters.Add(sgData.Cells[Lc, 0] + '=' + sgData.Cells[Lc, 1]);
      End;

      // Filtro especial para a coluna de Checkbox (Linha 1, Coluna 0)
      If sgData.Cells[0, 1] = '1' Then
        LFilters.Add('_CHECKED_=1');

      // Preservar posição do cursor se estivermos a editar a linha de filtros
      LSelStart := -1;
      LEditor := nil;
      If (sgData.Row = 1) AND Assigned(sgData.Editor) AND (sgData.Editor IS TCustomEdit) Then
      Begin
        LEditor := TCustomEdit(sgData.Editor);
        LSelStart := LEditor.SelStart;
        LSelLen := LEditor.SelLength;
      End;

      sgData.BeginUpdate;
      Try
        // Limpar apenas as linhas de dados, manter cabeçalho e filtros
        sgData.RowCount := 2;
      Finally
        sgData.EndUpdate;
      End;

      // Restaurar posição do cursor
      If (LSelStart >= 0) AND Assigned(LEditor) Then
      Begin
        LEditor.SelStart := LSelStart;
        LEditor.SelLength := LSelLen;
      End;

      lDataStaus.Caption := 'A filtrar e carregar...';
      pnlData.Caption := 'Dados';
      SplitterData.Top := 0;

      // Passar os filtros e as chaves marcadas para a thread
      FDataLoader := TDataLoaderThread.Create(Self, EtiquetaFile.Dados, sgData, LFilters, LCheckedKeys, NOT FDataLoadedOnce);
    Finally
      LFilters.Free;
      LCheckedKeys.Free;
    End;
  End;
End;

Procedure TfEtiqueta.UpdateObjectsFromData(ARow: Integer = -1);
Var
  Li, Lj, LBaseRow: Integer;
  LManager: TLabelManager;

  Procedure UpdateItem(AItem: TLabelItem);
  Var
    LTargetRow, LColIdx, Lc, Lr: Integer;
    LValue: String;
    LSubItem: TLabelItem;
  Begin
    If AItem.FieldName <> '' Then
    Begin
      LTargetRow := LBaseRow + AItem.RecordOffset;
      LValue := '';

      // Validar limites (linha 0 é cabeçalho, linha 1 é filtro, dados começam na 2)
      If (LTargetRow >= 2) AND (LTargetRow < sgData.RowCount) Then
      Begin
        LColIdx := -1;
        // Encontrar índice da coluna pelo nome (FieldName)
        For Lc := 1 To sgData.ColCount - 1 Do
        Begin
          If SameText(sgData.Cells[Lc, 0], AItem.FieldName) Then
          Begin
            LColIdx := Lc;
            Break;
          End;
        End;

        If LColIdx >= 0 Then
          LValue := sgData.Cells[LColIdx, LTargetRow];
      End;

      // Atualizar o objeto
      If AItem IS TLabelText Then
        TLabelText(AItem).Text := LValue
      Else If AItem IS TLabelBarcode Then
        TLabelBarcode(AItem).Text := LValue
      Else If AItem IS TLabelImage Then
      Begin
        If (LValue <> '') AND FileExists(LValue) Then
        Begin
          Try
            TLabelImage(AItem).Picture.LoadFromFile(LValue);
          Except
            TLabelImage(AItem).Picture.Clear;
          End;
        End
        Else
          TLabelImage(AItem).Picture.Clear;
      End;
    End;

    // Se for uma grelha, processar os filhos
    If AItem IS TLabelGrid Then
    Begin
      For Lr := 0 To TLabelGrid(AItem).Rows - 1 Do
        For Lc := 0 To TLabelGrid(AItem).Cols - 1 Do
        Begin
          LSubItem := TLabelGrid(AItem).GetCellItem(Lr, Lc);
          If Assigned(LSubItem) Then
            UpdateItem(LSubItem);
        End;
    End;
  End;

Begin
  If sgData.RowCount < 3 Then Exit; // Apenas cabeçalho e filtros

  LBaseRow := ARow;
  If LBaseRow < 0 Then LBaseRow := sgData.Row;
  If LBaseRow < 2 Then LBaseRow := 2;

  For Li := 0 To pcEtiquetas.PageCount - 1 Do
  Begin
    If (pcEtiquetas.Pages[Li].Tag = 0) Then Continue;
    LManager := TTabSettings(pcEtiquetas.Pages[Li].Tag).LabelManager;
    For Lj := 0 To LManager.Count - 1 Do
      UpdateItem(LManager[Lj]);
  End;

  // Invalidate all paintboxes
  For Li := 0 To pcEtiquetas.PageCount - 1 Do
  Begin
    If (pcEtiquetas.Pages[Li] <> tsMais) AND (pcEtiquetas.Pages[Li].ComponentCount > 0) Then
    Begin
      For Lj := 0 To pcEtiquetas.Pages[Li].ComponentCount - 1 Do
        If pcEtiquetas.Pages[Li].Components[Lj] IS TPaintBox Then
          TPaintBox(pcEtiquetas.Pages[Li].Components[Lj]).Invalidate;
    End;
    End;
  End;


Procedure TfEtiqueta.UpdatePropertiesPanel(AItem: TLabelItem);
Var
  LSelectedList: TList;
  Li, Lj: Integer;
  LCommonProps: TStringList;
  LPropName, LPropValue: String;
  LFirstItem: TLabelItem;
  LSameValue: Boolean;
  // Added: used when presenting propriedades do objeto contido numa célula
  LChildItem: TLabelItem;
  LP: TPoint;
  LHasCellObject: Boolean;
  LSG: TStringGrid;

  Procedure AddProp(Const AName, AValue: String);
  Var
    LRow: Integer;
  Begin
    LRow := LSG.RowCount;
    LSG.RowCount := LRow + 1;
    LSG.Cells[0, LRow] := AName;
    LSG.Cells[1, LRow] := AValue;
  End;

  Function GetPropValue(AObj: TLabelItem; Const APName: String): String;
  Var
    LGrid: TLabelGrid;
    LP: TPoint;
    LChildItem: TLabelItem;
    LRealPropName: String;
  Begin
    Result := '';
    LRealPropName := Trim(APName);

    // Se for uma propriedade da grelha (com prefixo)
    If Pos('(Grelha)', LRealPropName) > 0 Then
    Begin
      LRealPropName := Trim(StringReplace(LRealPropName, '(Grelha)', '', [rfReplaceAll]));
      // Chamada recursiva para o objeto grelha sem o prefixo
      Result := GetPropValue(AObj, LRealPropName);
      Exit;
    End;

    // Se estivermos numa grelha e houver um objeto na célula,
    // as propriedades sem prefixo referem-se ao objeto da célula.
    If (AObj IS TLabelGrid) Then
    Begin
      LGrid := TLabelGrid(AObj);
      If LGrid.GetSelectionCount = 1 Then
      Begin
        LP := LGrid.GetFirstSelectedCell;
        LChildItem := LGrid.GetCellItem(LP.Y, LP.X);
        If Assigned(LChildItem) Then
        Begin
          // Se for uma propriedade que o objeto filho possui, devolvemos o valor dele
          If SameText(LRealPropName, 'Nome') Then Result := LChildItem.Name
          Else If SameText(LRealPropName, 'Left (mm)') Then Result := FloatToStrF(PxToMm(LChildItem.Left), ffFixed, 7, 2)
          Else If SameText(LRealPropName, 'Top (mm)') Then Result := FloatToStrF(PxToMm(LChildItem.Top), ffFixed, 7, 2)
          Else If SameText(LRealPropName, 'Width (mm)') Then Result := FloatToStrF(PxToMm(LChildItem.Width), ffFixed, 7, 2)
          Else If SameText(LRealPropName, 'Height (mm)') Then Result := FloatToStrF(PxToMm(LChildItem.Height), ffFixed, 7, 2)
          Else If SameText(LRealPropName, 'Nome do campo') Then
          Begin
            Result := LChildItem.FieldName;
            If Result = '' Then Result := '<Vazio>';
          End
          Else If SameText(LRealPropName, 'Desvio de registos') Then Result := IntToStr(LChildItem.RecordOffset)
          Else If (LChildItem IS TLabelText) Then
          Begin
            If SameText(LRealPropName, 'Texto') Then Result := TLabelText(LChildItem).Text
            Else If SameText(LRealPropName, 'Font') Then Result := '' // Header
            Else If SameText(LRealPropName, 'Nome') Then Result := TLabelText(LChildItem).Font.Name
            Else If SameText(LRealPropName, 'Tamanho') Then Result := IntToStr(TLabelText(LChildItem).Font.Size)
            Else If SameText(LRealPropName, 'Cor') Then Result := ColorToString(TLabelText(LChildItem).Font.Color)
            Else If SameText(LRealPropName, 'Estilo') Then Result := '' // Header
            Else If SameText(LRealPropName, 'Negrito') Then Result := BoolToStr(fsBold IN TLabelText(LChildItem).Font.Style, True)
            Else If SameText(LRealPropName, 'Sublinhado') Then Result := BoolToStr(fsUnderline IN TLabelText(LChildItem).Font.Style, True)
            Else If SameText(LRealPropName, 'Itálico') Then Result := BoolToStr(fsItalic IN TLabelText(LChildItem).Font.Style, True)
            Else If SameText(LRealPropName, 'Rasurado') Then Result := BoolToStr(fsStrikeOut IN TLabelText(LChildItem).Font.Style, True)
            Else If SameText(LRealPropName, 'Alinham. H.') Then
            Begin
              Case TLabelText(LChildItem).Alignment Of
                taLeftJustify: Result := 'Esquerda';
                taCenter: Result := 'Centro';
                taRightJustify: Result := 'Direita';
              End;
            End
            Else If SameText(LRealPropName, 'Alinham. V.') Then
            Begin
              Case TLabelText(LChildItem).VertAlignment Of
                tlTop: Result := 'Topo';
                tlCenter: Result := 'Centro';
                tlBottom: Result := 'Base';
              End;
            End
            Else If SameText(LRealPropName, 'WordWrap') Then Result := BoolToStr(TLabelText(LChildItem).WordWrap, True);
          End
          Else If (LChildItem IS TLabelBarcode) Then
          Begin
            If SameText(LRealPropName, 'Texto') Then Result := TLabelBarcode(LChildItem).Text
            Else If SameText(LRealPropName, 'Mostrar Texto') Then Result := BoolToStr(TLabelBarcode(LChildItem).ShowText, True)
            Else If SameText(LRealPropName, 'Módulo') Then Result := FloatToStrF(TLabelBarcode(LChildItem).Module, ffFixed, 7, 2)
            Else If SameText(LRealPropName, 'Tipo') Then
            Begin
              Case TLabelBarcode(LChildItem).Kind Of
                bkCode128: Result := 'Code128';
                bkEAN13: Result := 'EAN13';
                bkQR: Result := 'QR';
              End;
            End
            Else If SameText(LRealPropName, 'Ratio') Then Result := FloatToStrF(TLabelBarcode(LChildItem).Ratio, ffFixed, 7, 2);
          End
          Else If (LChildItem IS TLabelImage) Then
          Begin
            If SameText(LRealPropName, 'Imagem') Then
            Begin
              If Assigned(TLabelImage(LChildItem).Picture) AND (TLabelImage(LChildItem).Picture.Graphic <> nil) AND (NOT TLabelImage(LChildItem).Picture.Graphic.Empty) Then
                Result := '(definida)'
              Else
                Result := '';
            End
            Else If SameText(LRealPropName, 'Ajustar') Then Result := BoolToStr(TLabelImage(LChildItem).Stretch, True)
            Else If SameText(LRealPropName, 'Proporcional') Then Result := BoolToStr(TLabelImage(LChildItem).Proportional, True);
          End;

          If Result <> '' Then Exit;
        End;
      End;
    End;

    // Propriedades padrão do objeto (ou da grelha se não houver objeto na célula)
    If LRealPropName = 'Nome' Then Result := AObj.Name
    Else If LRealPropName = 'Left (mm)' Then Result := FloatToStrF(PxToMm(AObj.Left), ffFixed, 7, 2)
    Else If LRealPropName = 'Top (mm)' Then Result := FloatToStrF(PxToMm(AObj.Top), ffFixed, 7, 2)
    Else If LRealPropName = 'Width (mm)' Then Result := FloatToStrF(PxToMm(AObj.Width), ffFixed, 7, 2)
    Else If LRealPropName = 'Height (mm)' Then Result := FloatToStrF(PxToMm(AObj.Height), ffFixed, 7, 2)
    Else If LRealPropName = 'Nome do campo' Then
    Begin
      Result := AObj.FieldName;
      If Result = '' Then Result := '<Vazio>';
    End
    Else If LRealPropName = 'Desvio de registos' Then Result := IntToStr(AObj.RecordOffset)
    Else If (AObj IS TLabelText) Then
    Begin
      If LRealPropName = 'Texto' Then Result := TLabelText(AObj).Text
      Else If LRealPropName = 'Font' Then Result := '' // Header
      Else If LRealPropName = 'Nome' Then Result := TLabelText(AObj).Font.Name
      Else If LRealPropName = 'Tamanho' Then Result := IntToStr(TLabelText(AObj).Font.Size)
      Else If LRealPropName = 'Cor' Then Result := ColorToString(TLabelText(AObj).Font.Color)
      Else If LRealPropName = 'Estilo' Then Result := '' // Header
      Else If LRealPropName = 'Negrito' Then Result := BoolToStr(fsBold IN TLabelText(AObj).Font.Style, True)
      Else If LRealPropName = 'Sublinhado' Then Result := BoolToStr(fsUnderline IN TLabelText(AObj).Font.Style, True)
      Else If LRealPropName = 'Itálico' Then Result := BoolToStr(fsItalic IN TLabelText(AObj).Font.Style, True)
      Else If LRealPropName = 'Rasurado' Then Result := BoolToStr(fsStrikeOut IN TLabelText(AObj).Font.Style, True)
      Else If LRealPropName = 'WordWrap' Then Result := BoolToStr(TLabelText(AObj).WordWrap, True)
      Else If LRealPropName = 'Alinham. H.' Then
      Begin
        Case TLabelText(AObj).Alignment Of
          taLeftJustify: Result := 'Esquerda';
          taCenter: Result := 'Centro';
          taRightJustify: Result := 'Direita';
        End;
      End
      Else If LRealPropName = 'Alinham. V.' Then
      Begin
        Case TLabelText(AObj).VertAlignment Of
          tlTop: Result := 'Topo';
          tlCenter: Result := 'Centro';
          tlBottom: Result := 'Base';
        End;
      End;
    End
    Else If (AObj IS TLabelBarcode) Then
    Begin
      If LRealPropName = 'Texto' Then Result := TLabelBarcode(AObj).Text
      Else If LRealPropName = 'Mostrar Texto' Then Result := BoolToStr(TLabelBarcode(AObj).ShowText, True)
      Else If LRealPropName = 'Módulo' Then Result := FloatToStrF(TLabelBarcode(AObj).Module, ffFixed, 7, 2)
      Else If LRealPropName = 'Ratio' Then Result := FloatToStrF(TLabelBarcode(AObj).Ratio, ffFixed, 7, 2)
      Else If LRealPropName = 'Tipo' Then
      Begin
        Case TLabelBarcode(AObj).Kind Of
          bkCode128: Result := 'Code128';
          bkEAN13: Result := 'EAN13';
          bkQR: Result := 'QR';
        End;
      End;
    End
    Else If (AObj IS TLabelImage) Then
    Begin
      If LRealPropName = 'Imagem' Then
      Begin
        If Assigned(TLabelImage(AObj).Picture) AND (TLabelImage(AObj).Picture.Graphic <> nil) AND (NOT TLabelImage(AObj).Picture.Graphic.Empty) Then
          Result := '(definida)'
        Else
          Result := '';
      End
      Else If LRealPropName = 'Ajustar' Then Result := BoolToStr(TLabelImage(AObj).Stretch, True)
      Else If LRealPropName = 'Proporcional' Then Result := BoolToStr(TLabelImage(AObj).Proportional, True);
    End
    Else If (AObj IS TLabelGrid) Then
    Begin
      LGrid := TLabelGrid(AObj);
      If LRealPropName = 'Linhas' Then Result := IntToStr(LGrid.Rows)
      Else If LRealPropName = 'Colunas' Then Result := IntToStr(LGrid.Cols)
      Else If LGrid.GetSelectionCount > 0 Then
      Begin
        LP := LGrid.GetFirstSelectedCell;
        If SameText(LRealPropName, 'Cor Fundo') Then Result := ColorToString(LGrid.GetCellBackColor(LP.Y, LP.X))
        Else If SameText(LRealPropName, 'Cor Borda') Then Result := ColorToString(LGrid.GetCellBorderColor(LP.Y, LP.X))
        Else If SameText(LRealPropName, 'Esquerda') Then Result := BoolToStr(beLeft IN LGrid.GetCellBorderVisible(LP.Y, LP.X), True)
        Else If SameText(LRealPropName, 'Topo') Then Result := BoolToStr(beTop IN LGrid.GetCellBorderVisible(LP.Y, LP.X), True)
        Else If SameText(LRealPropName, 'Direita') Then Result := BoolToStr(beRight IN LGrid.GetCellBorderVisible(LP.Y, LP.X), True)
        Else If SameText(LRealPropName, 'Baixo') Then Result := BoolToStr(beBottom IN LGrid.GetCellBorderVisible(LP.Y, LP.X), True);
      End;
    End;
  End;

Begin
  LSG := GetActivePropertiesGrid;
  If LSG = nil Then Exit;

  If LabelManager = nil Then Exit;

  LSelectedList := LabelManager.GetSelectedItems;
  Try
    If LSelectedList.Count = 0 Then
    Begin
      ShowLabelProperties;
      Exit;
    End;

    LSG.Enabled := True;
    LSG.RowCount := 1; // Reset

    // Define potential common properties
    LCommonProps := TStringList.Create;
    Try
      LCommonProps.Add('Left (mm)');
      LCommonProps.Add('Top (mm)');
      LCommonProps.Add('Width (mm)');
      LCommonProps.Add('Height (mm)');

      LCommonProps.Add('Nome do campo');
      LCommonProps.Add('Desvio de registos');

      LFirstItem := TLabelItem(LSelectedList[0]);

      If LSelectedList.Count = 1 Then
      Begin
        LCommonProps.Insert(0, 'Nome');
        If LFirstItem IS TLabelText Then
        Begin
          LCommonProps.Add('Texto');

          If FFontExpanded Then
          Begin
            LCommonProps.Add(sMenos + 'Font');
            LCommonProps.Add(sPropIndent + 'Nome');
            LCommonProps.Add(sPropIndent + 'Tamanho');
            LCommonProps.Add(sPropIndent + 'Cor');

            If FFontStyleExpanded Then
            Begin
              LCommonProps.Add(sPropIndent + sMenos + 'Estilo');
              LCommonProps.Add(sPropIndent + sPropIndent + 'Negrito');
              LCommonProps.Add(sPropIndent + sPropIndent + 'Sublinhado');
              LCommonProps.Add(sPropIndent + sPropIndent + 'Itálico');
              LCommonProps.Add(sPropIndent + sPropIndent + 'Rasurado');
            End
            Else
              LCommonProps.Add(sPropIndent + sMais + 'Estilo');
          End
          Else
            LCommonProps.Add(sMais + 'Font');

          LCommonProps.Add('Alinham. H.');
          LCommonProps.Add('Alinham. V.');
          LCommonProps.Add('WordWrap');
        End
        Else If LFirstItem IS TLabelBarcode Then
        Begin
          LCommonProps.Add('Texto');
          LCommonProps.Add('Tipo');
          LCommonProps.Add('Mostrar Texto');
          LCommonProps.Add('Módulo');
          LCommonProps.Add('Ratio');
        End
        Else If LFirstItem IS TLabelImage Then
        Begin
          LCommonProps.Add('Imagem');
          LCommonProps.Add('Ajustar');
          LCommonProps.Add('Proporcional');
        End
        Else If LFirstItem IS TLabelGrid Then
        Begin
          // Por omissão, assumimos que não há objeto em célula selecionada
          LHasCellObject := False;

          If TLabelGrid(LFirstItem).GetSelectionCount > 0 Then
          Begin
            If TLabelGrid(LFirstItem).GetSelectionCount = 1 Then
            Begin
              LP := TLabelGrid(LFirstItem).GetFirstSelectedCell;
              LChildItem := TLabelGrid(LFirstItem).GetCellItem(LP.Y, LP.X);
              If Assigned(LChildItem) Then
              Begin
                LHasCellObject := True;
                // Propriedades do Objeto (Principais) - Já adicionadas no início, mas aqui definimos o contexto

                If LChildItem IS TLabelText Then
                Begin
                  LCommonProps.Add('Texto');
                  LCommonProps.Add('Font');
                  LCommonProps.Add('Tamanho');
                  LCommonProps.Add('Cor');
                  LCommonProps.Add('WordWrap');
                  LCommonProps.Add('Alinham. H.');
                  LCommonProps.Add('Alinham. V.');
                End
                Else If LChildItem IS TLabelBarcode Then
                Begin
                  LCommonProps.Add('Texto');
                  LCommonProps.Add('Tipo');
                  LCommonProps.Add('Mostrar Texto');
                  LCommonProps.Add('Módulo');
                  LCommonProps.Add('Ratio');
                End
                Else If LChildItem IS TLabelImage Then
                Begin
                  LCommonProps.Add('Imagem');
                  LCommonProps.Add('Ajustar');
                  LCommonProps.Add('Proporcional');
                End;
              End;
            End;

            // Grupo Célula (Secundário)
            If FCellPropsExpanded Then
              LCommonProps.Add(sMenos + 'Célula')
            Else
              LCommonProps.Add(sMais + 'Célula');

            If FCellPropsExpanded Then
            Begin
              LCommonProps.Add(sPropIndent + 'Cor Fundo');
              LCommonProps.Add(sPropIndent + 'Cor Borda');

              If FBordersExpanded Then
                LCommonProps.Add(sPropIndent + sMenos + 'Bordas')
              Else
                LCommonProps.Add(sPropIndent + sMais + 'Bordas');

              If FBordersExpanded Then
              Begin
                LCommonProps.Add(sPropIndent + sPropIndent + 'Esquerda');
                LCommonProps.Add(sPropIndent + sPropIndent + 'Topo');
                LCommonProps.Add(sPropIndent + sPropIndent + 'Direita');
                LCommonProps.Add(sPropIndent + sPropIndent + 'Baixo');
              End;
            End;

            // Grupo Grelha (Terciário)
            If LHasCellObject Then
            Begin
              If FGridPropsExpanded Then
                LCommonProps.Add(sMenos + 'Grelha')
              Else
                LCommonProps.Add(sMais + 'Grelha');

              If FGridPropsExpanded Then
              Begin
                LCommonProps.Add(sPropIndent + '(Grelha) Nome');
                LCommonProps.Add(sPropIndent + '(Grelha) Linhas');
                LCommonProps.Add(sPropIndent + '(Grelha) Colunas');
                LCommonProps.Add(sPropIndent + '(Grelha) Left (mm)');
                LCommonProps.Add(sPropIndent + '(Grelha) Top (mm)');
              End;
            End
            Else
            Begin
              LCommonProps.Add('Linhas');
              LCommonProps.Add('Colunas');
            End;
          End;
        End;
      End
      Else
      Begin
        // Multiple items: Check if they are all same type

        // Check if all are Text
        LSameValue := True;
        For Li := 1 To LSelectedList.Count - 1 Do
          If NOT (TLabelItem(LSelectedList[Li]) IS TLabelText) Then LSameValue := False;

        If LSameValue AND (LFirstItem IS TLabelText) Then
        Begin
          LCommonProps.Add('Texto');
          LCommonProps.Add('Font');
          LCommonProps.Add('Tamanho');
          LCommonProps.Add('Cor');
          LCommonProps.Add('Alinham. H.');
          LCommonProps.Add('Alinham. V.');
        End;

        // Check if all are Barcode
        LSameValue := True;
        For Li := 1 To LSelectedList.Count - 1 Do
          If NOT (TLabelItem(LSelectedList[Li]) IS TLabelBarcode) Then LSameValue := False;

        If LSameValue AND (LFirstItem IS TLabelBarcode) Then
        Begin
          LCommonProps.Add('Texto');
          LCommonProps.Add('Tipo');
          LCommonProps.Add('Mostrar Texto');
          LCommonProps.Add('Módulo');
        End;

        // Check if all are Grid
        LSameValue := True;
        For Li := 1 To LSelectedList.Count - 1 Do
          If NOT (TLabelItem(LSelectedList[Li]) IS TLabelGrid) Then LSameValue := False;

        If LSameValue AND (LFirstItem IS TLabelGrid) Then
        Begin
          LCommonProps.Add('Linhas');
          LCommonProps.Add('Colunas');
        End;
      End;

      // Now iterate properties and check values
      For Li := 0 To LCommonProps.Count - 1 Do
      Begin
        LPropName := LCommonProps[Li];
        LPropValue := GetPropValue(LFirstItem, LPropName);
        LSameValue := True;

        For Lj := 1 To LSelectedList.Count - 1 Do
        Begin
          If GetPropValue(TLabelItem(LSelectedList[Lj]), LPropName) <> LPropValue Then
          Begin
            LSameValue := False;
            Break;
          End;
        End;

        If LSameValue Then
          AddProp(LPropName, LPropValue)
        Else
          AddProp(LPropName, ''); // Empty for different values
      End;

    Finally
      LCommonProps.Free;
    End;

    pObjetPlusResize(nil);

    // Update Equalize Buttons State
    sbEqualizeWidth.Enabled := False;
    sbEqualizeHeight.Enabled := False;

    If LSelectedList.Count > 1 Then
    Begin
      sbEqualizeWidth.Enabled := True;
      sbEqualizeHeight.Enabled := True;
    End
    Else If LSelectedList.Count = 1 Then
    Begin
      If TLabelItem(LSelectedList[0]) IS TLabelGrid Then
      Begin
        // Check for multiple selected cells in grid
        // This is a simplified check, ideally we check for >1 cols or >1 rows specifically
        If TLabelGrid(LSelectedList[0]).GetSelectionCount > 1 Then
        Begin
          sbEqualizeWidth.Enabled := True;
          sbEqualizeHeight.Enabled := True;
        End;
      End;
    End;

  Finally
    LSelectedList.Free;
  End;
End;

Function TfEtiqueta.IsPaperSizeEditable(Const APaperName: String): Boolean;
Var
  LUName: String;
Begin
  LUName := UpperCase(APaperName);
  // Tamanhos personalizados ou definidos pelo utilizador são sempre editáveis
  If (Pos('CUSTOM', LUName) > 0) OR (Pos('USER', LUName) > 0) OR (Pos('PERSONALIZADO', LUName) > 0) OR (Pos('DEFINIDO PELO UTILIZADOR', LUName) > 0) OR (APaperName = '') Then
  Begin
    Result := True;
    Exit;
  End;

  // Tamanhos de papel comuns (Séries A, B, Letter, etc.)
  // Estes permitem ajustes pois os fabricantes de papel podem ter ligeiras variações
  Result := (Pos('A0', LUName) > 0) OR (Pos('A1', LUName) > 0) OR (Pos('A2', LUName) > 0) OR (Pos('A3', LUName) > 0) OR (Pos('A4', LUName) > 0) OR
    (Pos('A5', LUName) > 0) OR (Pos('A6', LUName) > 0) OR (Pos('B0', LUName) > 0) OR (Pos('B1', LUName) > 0) OR (Pos('B2', LUName) > 0) OR
    (Pos('B3', LUName) > 0) OR (Pos('B4', LUName) > 0) OR (Pos('B5', LUName) > 0) OR (Pos('LETTER', LUName) > 0) OR (Pos('LEGAL', LUName) > 0) OR
    (Pos('EXECUTIVE', LUName) > 0) OR (Pos('STATEMENT', LUName) > 0) OR (Pos('FOLIO', LUName) > 0) OR (Pos('LEDGER', LUName) > 0) OR
    (Pos('TABLOID', LUName) > 0) OR (Pos('ENVELOPE', LUName) > 0);
End;

Function TfEtiqueta.IsPaperSizeContinuous(Const APaperName: String): Boolean;
Var
  LUName: String;
  Li, LNumCount: Integer;
  LInNum: Boolean;
Begin
  LUName := UpperCase(APaperName);

  // 1. Etiquetas circulares NÃƒO são contínuas (têm diâmetro fixo)
  If (Pos('DIA', LUName) > 0) OR (Pos('DIAMETER', LUName) > 0) Then
  Begin
    Result := False;
    Exit;
  End;

  // 2. Verificação por palavras-chave conhecidas
  If (Pos('CONTINUOUS', LUName) > 0) OR (Pos('CONTÃNUO', LUName) > 0) OR (Pos('ROLL', LUName) > 0) OR (Pos('ROLO', LUName) > 0) Then
  Begin
    Result := True;
    Exit;
  End;

  // 3. Heurística: Apenas uma dimensão mencionada (ex: "62mm" vs "62mm x 29mm")
  LNumCount := 0;
  LInNum := False;
  For Li := 1 To Length(LUName) Do
  Begin
    If LUName[Li] IN ['0'..'9'] Then
    Begin
      If NOT LInNum Then
      Begin
        Inc(LNumCount);
        LInNum := True;
      End;
    End
    Else
      LInNum := False;
  End;

  // Se tem exatamente um número e menciona 'mm', e começa por número
  // é um rolo contínuo.
  Result := (LNumCount = 1) AND (Pos('MM', LUName) > 0) AND (LUName[1] IN ['0'..'9']);
End;

Function TfEtiqueta.PxToMm(APx: Integer): Double;
Begin
  Result := APx * 25.4 / Screen.PixelsPerInch;
End;

Function TfEtiqueta.MmToPx(AMm: Double): Integer;
Begin
  Result := Round(AMm * Screen.PixelsPerInch / 25.4);
End;

Procedure TfEtiqueta.sgPropertiesEditorKeyPress(ASender: TObject; Var AKey: Char);
Var
  LPropName, LS: String;
  LDecSep: Char;
  LSG: TStringGrid;
Begin
  LSG := GetActivePropertiesGrid;
  If LSG = nil Then Exit;

  LPropName := LSG.Cells[0, LSG.Row];

  // List of numeric properties (Integer or Float)
  If (Pos('mm', LPropName) > 0) OR (Trim(LPropName) = 'Tamanho') OR (LPropName = 'Módulo') OR (LPropName = 'Ratio') Then
  Begin
    LDecSep := DefaultFormatSettings.DecimalSeparator;

    If (AKey = '.') OR (AKey = ',') Then AKey := LDecSep;

    If NOT (AKey IN ['0'..'9', #8, LDecSep]) Then
    Begin
      AKey := #0;
      Exit;
    End;

    If (AKey = LDecSep) Then
    Begin
      LS := TEdit(ASender).Text;
      If Pos(LDecSep, LS) > 0 Then AKey := #0;
    End;
  End;
End;

Procedure TfEtiqueta.sgPropertiesSelectCell(ASender: TObject; ACol, ARow: Integer; Var ACanSelect: Boolean);
Var
  LPropName: String;
  LR: TRect;
  Li: Integer;
  LSG: TStringGrid;
Begin
  // Guard clauses
  If (csLoading IN ComponentState) OR (csDestroying IN ComponentState) Then Exit;
  LSG := ASender AS TStringGrid;

  If (ARow < 0) OR (ARow >= LSG.RowCount) Then Exit;
  If (ACol < 0) OR (ACol >= LSG.ColCount) Then Exit;

  LPropName := LSG.Cells[0, ARow];

  // Prevent selection and editing of parent nodes (headers)
  If (Pos(sMais, LPropName) > 0) OR (Pos(sMenos, LPropName) > 0) Then
  Begin
    ACanSelect := False;
    Exit;
  End;

  If ACol = 1 Then
  Begin
    LSG.Options := LSG.Options + [goEditing];

    // Reset editors
    If (LSG.Columns.Count > 1) Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsAuto;
      If Assigned(LSG.Columns[1].PickList) Then
        LSG.Columns[1].PickList.Clear;
    End;

    // Configure editor based on property name
    If SameText(Trim(LPropName), 'Nome') Then
    Begin
      // Se tiver indentação, é o Nome da Font
      If Pos(sPropIndent, LPropName) > 0 Then
      Begin
        LSG.Columns[1].ButtonStyle := cbsPickList;
        If (LSG.Columns.Count > 1) AND Assigned(LSG.Columns[1].PickList) Then
          LSG.Columns[1].PickList.Assign(Screen.Fonts);
      End;
    End
    Else If SameText(Trim(LPropName), 'Negrito') OR SameText(Trim(LPropName), 'Sublinhado') OR SameText(Trim(LPropName), 'Itálico') OR
      SameText(Trim(LPropName), 'Rasurado') Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsPickList;
      If (LSG.Columns.Count > 1) AND Assigned(LSG.Columns[1].PickList) Then
      Begin
        LSG.Columns[1].PickList.Add('True');
        LSG.Columns[1].PickList.Add('False');
      End;
    End
    Else If SameText(Trim(LPropName), 'Tipo') Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsPickList;
      If (LSG.Columns.Count > 1) AND Assigned(LSG.Columns[1].PickList) Then
      Begin
        LSG.Columns[1].PickList.Add('Code128');
        LSG.Columns[1].PickList.Add('EAN13');
        LSG.Columns[1].PickList.Add('QR');
      End;
    End
    Else If SameText(Trim(LPropName), 'Mostrar Texto') OR SameText(Trim(LPropName), 'Esquerda') OR SameText(Trim(LPropName), 'Topo') OR
      SameText(Trim(LPropName), 'Direita') OR SameText(Trim(LPropName), 'Baixo') Then
    Begin
      If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
      Begin
        // Boolean property for Grid - Prevent editing (we handle click)
        LSG.Options := LSG.Options - [goEditing];
      End
      Else If SameText(Trim(LPropName), 'Mostrar Texto') AND (FSelectedItem <> nil) AND (FSelectedItem IS TLabelBarcode) Then
      Begin
        // Boolean property for Barcode - Prevent editing
        LSG.Options := LSG.Options - [goEditing];
      End;
    End
    Else If SameText(Trim(LPropName), 'Impressora') Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsPickList;
      If (LSG.Columns.Count > 1) AND Assigned(LSG.Columns[1].PickList) Then
        LSG.Columns[1].PickList.Assign(Printer.Printers);
    End
    Else If SameText(Trim(LPropName), 'Papel') Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsPickList;
      If (LSG.Columns.Count > 1) AND Assigned(LSG.Columns[1].PickList) Then
      Begin
        If GetActiveSettings <> nil Then
          Printer.SetPrinter(GetActiveSettings.PrinterName);
        LSG.Columns[1].PickList.Assign(Printer.PaperSize.SupportedPapers);
      End;
    End
    Else If SameText(Trim(LPropName), 'Orientação') Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsPickList;
      If (LSG.Columns.Count > 1) AND Assigned(LSG.Columns[1].PickList) Then
      Begin
        LSG.Columns[1].PickList.Add('Portrait');
        LSG.Columns[1].PickList.Add('Landscape');
      End;
    End
    Else If SameText(Trim(LPropName), 'Alinham. H.') Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsPickList;
      If (LSG.Columns.Count > 1) AND Assigned(LSG.Columns[1].PickList) Then
      Begin
        LSG.Columns[1].PickList.Add('Esquerda');
        LSG.Columns[1].PickList.Add('Centro');
        LSG.Columns[1].PickList.Add('Direita');
      End;
    End
    Else If SameText(Trim(LPropName), 'Alinham. V.') Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsPickList;
      If (LSG.Columns.Count > 1) AND Assigned(LSG.Columns[1].PickList) Then
      Begin
        LSG.Columns[1].PickList.Add('Topo');
        LSG.Columns[1].PickList.Add('Centro');
        LSG.Columns[1].PickList.Add('Base');
      End;
    End
    Else If SameText(Trim(LPropName), 'Imagem') OR SameText(Trim(LPropName), '(Objeto) Imagem') Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsEllipsis;
    End
    Else If SameText(Trim(LPropName), 'Nome do campo') Then
    Begin
      LSG.Columns[1].ButtonStyle := cbsPickList;
      If (LSG.Columns.Count > 1) AND Assigned(LSG.Columns[1].PickList) Then
      Begin
        LSG.Columns[1].PickList.Add('<Vazio>');
        For Li := 1 To sgData.ColCount - 1 Do
          LSG.Columns[1].PickList.Add(sgData.Cells[Li, 0]);
      End;
    End
    Else If SameText(Trim(LPropName), 'Cor') OR SameText(Trim(LPropName), 'Cor Fundo') OR SameText(Trim(LPropName), 'Cor Borda') Then
    Begin
      // Position and show ColorBox
      LR := LSG.CellRect(ACol, ARow);
      cbColorEditor0.SetBounds(LSG.Left + LR.Left, LSG.Top + LR.Top, LR.Right - LR.Left, LR.Bottom - LR.Top);
      cbColorEditor0.Selected := StringToColorDef(LSG.Cells[ACol, ARow], clBlack);
      cbColorEditor0.Visible := True;
      cbColorEditor0.BringToFront;
      cbColorEditor0.SetFocus;
      // Disable grid editing for this cell to avoid conflict
      LSG.Options := LSG.Options - [goEditing];
    End
    Else If SameText(Trim(LPropName), 'Largura (mm)') Then
    Begin
      // In continuous paper, width is fixed by the roll, but height is variable
      If (GetActiveSettings <> nil) AND (IsPaperSizeContinuous(GetActiveSettings.PaperName) OR NOT IsPaperSizeEditable(GetActiveSettings.PaperName)) Then
        LSG.Options := LSG.Options - [goEditing];
    End
    Else If SameText(Trim(LPropName), 'Altura (mm)') Then
    Begin
      // In continuous paper, height is always editable
      If (GetActiveSettings <> nil) AND (NOT IsPaperSizeContinuous(GetActiveSettings.PaperName) AND NOT IsPaperSizeEditable(GetActiveSettings.PaperName)) Then
        LSG.Options := LSG.Options - [goEditing];
    End;

    // Attach KeyPress filter to editor
    If (LSG.Options * [goEditing] <> []) AND Assigned(LSG.Editor) Then
      TEdit(LSG.Editor).OnKeyPress := sgPropertiesEditorKeyPress;
  End
  Else
  Begin
    LSG.Options := LSG.Options - [goEditing];
    cbColorEditor0.Visible := False;
  End;
End;

Procedure TfEtiqueta.sgPropertiesSelectEditor(ASender: TObject; ACol, ARow: Integer; Var AEditor: TWinControl);
Begin
  If (AEditor IS TCustomComboBox) Then
    TCustomComboBox(AEditor).Style := csDropDownList;
End;

Procedure TfEtiqueta.sgPropertiesButtonClick(ASender: TObject; ACol, ARow: Integer);
Var
  LPropName: String;
  LSelectedList: TList;
  Li: Integer;
  LItem, LChildItem: TLabelItem;
  LGrid: TLabelGrid;
  LP: TPoint;
  LSG: TStringGrid;
Begin
  LSG := ASender AS TStringGrid;
  LPropName := LSG.Cells[0, LSG.Row];

  If SameText(Trim(LPropName), 'Imagem') OR SameText(Trim(LPropName), '(Objeto) Imagem') Then
  Begin
    If opdImagem.Execute Then
    Begin
      If SameText(Trim(LPropName), '(Objeto) Imagem') Then
      Begin
        // Handle child object in grid
        If (FSelectedItem IS TLabelGrid) Then
        Begin
          LGrid := TLabelGrid(FSelectedItem);
          If LGrid.GetSelectionCount = 1 Then
          Begin
            LP := LGrid.GetFirstSelectedCell;
            LChildItem := LGrid.GetCellItem(LP.Y, LP.X);
            If (LChildItem IS TLabelImage) Then
            Begin
              Try
                TLabelImage(LChildItem).Picture.LoadFromFile(opdImagem.FileName);
              Except
                On E: Exception Do
                  ShowMessage('Erro ao carregar imagem: ' + E.Message);
              End;
            End;
          End;
        End;
      End
      Else
      Begin
        // Handle normal selected images
        LSelectedList := LabelManager.GetSelectedItems;
        Try
          For Li := 0 To LSelectedList.Count - 1 Do
          Begin
            LItem := TLabelItem(LSelectedList[Li]);
            If LItem IS TLabelImage Then
            Begin
              Try
                TLabelImage(LItem).Picture.LoadFromFile(opdImagem.FileName);
              Except
                On E: Exception Do
                  ShowMessage('Erro ao carregar imagem: ' + E.Message);
              End;
            End;
          End;
        Finally
          LSelectedList.Free;
        End;
      End;
      GetActivePaintBox.Invalidate;
      UpdatePropertiesPanel(FSelectedItem);
    End;
  End;
End;

Procedure TfEtiqueta.cbColorEditorChange(ASender: TObject);
Var
  LSG: TStringGrid;
Begin
  LSG := GetActivePropertiesGrid;
  If LSG = nil Then Exit;

  If (LSG.Row > 0) AND (LSG.Row < LSG.RowCount) Then
  Begin
    LSG.Cells[1, LSG.Row] := ColorToString(cbColorEditor0.Selected);
    sgPropertiesSetEditText(LSG, 1, LSG.Row, ColorToString(cbColorEditor0.Selected));
  End;
  cbColorEditor0.Visible := False;
  If LSG.CanFocus Then LSG.SetFocus;
End;

Procedure TfEtiqueta.cbColorEditorExit(ASender: TObject);
Begin
  cbColorEditor0.Visible := False;
End;

Procedure TfEtiqueta.sgPropertiesSetEditText(ASender: TObject; ACol, ARow: Integer; Const AValue: String);
Var
  LPropName: String;
  LSelectedList: TList;
  Li, Lr, Lc: Integer;
  LItem, LChildItem: TLabelItem;
  LGrid: TLabelGrid;
  LP: TPoint;
  LRealPropName: String;
  LSG: TStringGrid;
  LSettings: TTabSettings;
Begin
  LSettings := GetActiveSettings;
  If LSettings = nil Then Exit;

  LSG := ASender AS TStringGrid;

  LSelectedList := LabelManager.GetSelectedItems;
  Try
    If LSelectedList.Count = 0 Then
    Begin
      LPropName := Trim(LSG.Cells[0, ARow]);
      If LPropName = 'Impressora' Then
      Begin
        If AValue <> LSettings.PrinterName Then
        Begin
          LSettings.PrinterName := AValue;
          LSettings.PaperWidth := 0;
          LSettings.PaperHeight := 0;
          Try
            Printer.SetPrinter(LSettings.PrinterName);
            LSettings.PaperName := Printer.PaperSize.PaperName;
            UpdateLabelSize;
            ShowLabelProperties;
          Except
            SysUtils.Beep;
          End;
        End;
      End
      Else If LPropName = 'Papel' Then
      Begin
        If AValue <> LSettings.PaperName Then
        Begin
          LSettings.PaperName := AValue;
          LSettings.PaperWidth := 0;
          LSettings.PaperHeight := 0;
          UpdateLabelSize;
          ShowLabelProperties;
        End;
      End
      Else If SameText(Trim(LPropName), 'Esquerda (mm)') Then
      Begin
        LSettings.Margins.Left := MmToPx(StrToFloatDef(AValue, PxToMm(LSettings.Margins.Left)));
        GetActivePaintBox.Invalidate;
      End
      Else If SameText(Trim(LPropName), 'Topo (mm)') Then
      Begin
        LSettings.Margins.Top := MmToPx(StrToFloatDef(AValue, PxToMm(LSettings.Margins.Top)));
        GetActivePaintBox.Invalidate;
      End
      Else If SameText(Trim(LPropName), 'Direita (mm)') Then
      Begin
        LSettings.Margins.Right := MmToPx(StrToFloatDef(AValue, PxToMm(LSettings.Margins.Right)));
        GetActivePaintBox.Invalidate;
      End
      Else If SameText(Trim(LPropName), 'Baixo (mm)') Then
      Begin
        LSettings.Margins.Bottom := MmToPx(StrToFloatDef(AValue, PxToMm(LSettings.Margins.Bottom)));
        GetActivePaintBox.Invalidate;
      End
      Else If LPropName = 'Orientação' Then
      Begin
        If SameText(AValue, 'Portrait') Then
          LSettings.Orientation := poPortrait
        Else If SameText(AValue, 'Landscape') Then
          LSettings.Orientation := poLandscape;
        UpdateLabelSize;
        ShowLabelProperties;
      End
      Else If SameText(Trim(LPropName), 'Colunas') Then Begin LSettings.LayoutCols := StrToIntDef(AValue, 1); GetActivePaintBox.Invalidate; End
      Else If SameText(Trim(LPropName), 'Linhas') Then Begin LSettings.LayoutRows := StrToIntDef(AValue, 1); GetActivePaintBox.Invalidate; End
      Else If SameText(Trim(LPropName), 'Espaço X (mm)') Then Begin LSettings.LayoutGapX := MmToPx(StrToFloatDef(AValue, 0)); GetActivePaintBox.Invalidate; End
      Else If SameText(Trim(LPropName), 'Espaço Y (mm)') Then Begin LSettings.LayoutGapY := MmToPx(StrToFloatDef(AValue, 0)); GetActivePaintBox.Invalidate; End
      Else If SameText(Trim(LPropName), 'Largura (mm)') Then
      Begin
        LSettings.PaperWidth := MmToPx(StrToFloatDef(AValue, PxToMm(LSettings.PaperWidth)));
        UpdateLabelSize;
      End
      Else If SameText(Trim(LPropName), 'Altura (mm)') Then
      Begin
        LSettings.PaperHeight := MmToPx(StrToFloatDef(AValue, PxToMm(LSettings.PaperHeight)));
        UpdateLabelSize;
      End;
      Exit;
    End;

    If ACol <> 1 Then Exit;
    If (ARow < 0) OR (ARow >= LSG.RowCount) Then Exit;

    LPropName := Trim(LSG.Cells[0, ARow]);

    Try
      For Li := 0 To LSelectedList.Count - 1 Do
      Begin
        LItem := TLabelItem(LSelectedList[Li]);

        // Se for uma propriedade da grelha (com prefixo)
        If Pos('(Grelha)', LPropName) > 0 Then
        Begin
          LRealPropName := Trim(StringReplace(LPropName, '(Grelha)', '', [rfReplaceAll]));
          If SameText(LRealPropName, 'Nome') Then
          Begin
            If LabelManager.ExistsName(AValue, LItem) Then
            Begin
              ShowMessage('Já existe um objeto com o nome "' + AValue + '".');
              Exit;
            End;
            LItem.Name := AValue;
          End
          Else If SameText(LRealPropName, 'Left (mm)') Then LItem.Left := MmToPx(StrToFloatDef(AValue, PxToMm(LItem.Left)))
          Else If SameText(LRealPropName, 'Top (mm)') Then LItem.Top := MmToPx(StrToFloatDef(AValue, PxToMm(LItem.Top)))
          Else If SameText(LRealPropName, 'Linhas') AND (LItem IS TLabelGrid) Then TLabelGrid(LItem).Rows := StrToIntDef(AValue, TLabelGrid(LItem).Rows)
          Else If SameText(LRealPropName, 'Colunas') AND (LItem IS TLabelGrid) Then TLabelGrid(LItem).Cols := StrToIntDef(AValue, TLabelGrid(LItem).Cols);
          Continue;
        End;

        If (LSG.Cells[0, ARow] = 'Nome') Then
        Begin
          If LabelManager.ExistsName(AValue, LItem) Then
          Begin
            ShowMessage('Já existe um objeto com o nome "' + AValue + '".');
            Exit;
          End;
          LItem.Name := AValue;
        End
        Else If SameText(Trim(LPropName), 'Left (mm)') Then LItem.Left := MmToPx(StrToFloatDef(AValue, PxToMm(LItem.Left)))
        Else If SameText(Trim(LPropName), 'Top (mm)') Then LItem.Top := MmToPx(StrToFloatDef(AValue, PxToMm(LItem.Top)))
        Else If SameText(Trim(LPropName), 'Width (mm)') Then LItem.Width := MmToPx(StrToFloatDef(AValue, PxToMm(LItem.Width)))
        Else If SameText(Trim(LPropName), 'Height (mm)') Then LItem.Height := MmToPx(StrToFloatDef(AValue, PxToMm(LItem.Height)))
        Else If SameText(Trim(LPropName), 'Nome do campo') Then
        Begin
          If SameText(AValue, '<Vazio>') Then
            LItem.FieldName := ''
          Else
            LItem.FieldName := AValue;
        End
        Else If SameText(Trim(LPropName), 'Desvio de registos') Then LItem.RecordOffset := StrToIntDef(AValue, LItem.RecordOffset);

        If LItem IS TLabelText Then
        Begin
          If SameText(Trim(LPropName), 'Texto') Then TLabelText(LItem).Text := AValue
          Else If SameText(Trim(LPropName), 'Font') Then TLabelText(LItem).Font.Name := AValue // Fallback legacy
          Else If SameText(Trim(LPropName), 'Nome') Then TLabelText(LItem).Font.Name := AValue
          Else If SameText(Trim(LPropName), 'Tamanho') Then TLabelText(LItem).Font.Size := StrToIntDef(AValue, TLabelText(LItem).Font.Size)
          Else If SameText(Trim(LPropName), 'Cor') Then TLabelText(LItem).Font.Color := StringToColorDef(AValue, TLabelText(LItem).Font.Color)

          // Styles
          Else If SameText(Trim(LPropName), 'Negrito') Then
          Begin
            If StrToBoolDef(AValue, False) Then TLabelText(LItem).Font.Style := TLabelText(LItem).Font.Style + [fsBold]
            Else
              TLabelText(LItem).Font.Style := TLabelText(LItem).Font.Style - [fsBold];
          End
          Else If SameText(Trim(LPropName), 'Sublinhado') Then
          Begin
            If StrToBoolDef(AValue, False) Then TLabelText(LItem).Font.Style := TLabelText(LItem).Font.Style + [fsUnderline]
            Else
              TLabelText(LItem).Font.Style := TLabelText(LItem).Font.Style - [fsUnderline];
          End
          Else If SameText(Trim(LPropName), 'Itálico') Then
          Begin
            If StrToBoolDef(AValue, False) Then TLabelText(LItem).Font.Style := TLabelText(LItem).Font.Style + [fsItalic]
            Else
              TLabelText(LItem).Font.Style := TLabelText(LItem).Font.Style - [fsItalic];
          End
          Else If SameText(Trim(LPropName), 'Rasurado') Then
          Begin
            If StrToBoolDef(AValue, False) Then TLabelText(LItem).Font.Style := TLabelText(LItem).Font.Style + [fsStrikeOut]
            Else
              TLabelText(LItem).Font.Style := TLabelText(LItem).Font.Style - [fsStrikeOut];
          End

          Else If SameText(Trim(LPropName), 'WordWrap') Then TLabelText(LItem).WordWrap := StrToBoolDef(AValue, TLabelText(LItem).WordWrap)
          Else If SameText(Trim(LPropName), 'Alinham. H.') Then
          Begin
            If SameText(AValue, 'Esquerda') Then TLabelText(LItem).Alignment := taLeftJustify
            Else If SameText(AValue, 'Centro') Then TLabelText(LItem).Alignment := taCenter
            Else If SameText(AValue, 'Direita') Then TLabelText(LItem).Alignment := taRightJustify;
          End
          Else If SameText(Trim(LPropName), 'Alinham. V.') Then
          Begin
            If SameText(AValue, 'Topo') Then TLabelText(LItem).VertAlignment := tlTop
            Else If SameText(AValue, 'Centro') Then TLabelText(LItem).VertAlignment := tlCenter
            Else If SameText(AValue, 'Base') Then TLabelText(LItem).VertAlignment := tlBottom;
          End;
        End
        Else If LItem IS TLabelBarcode Then
        Begin
          If SameText(Trim(LPropName), 'Texto') Then
            TLabelBarcode(LItem).Text := AValue
          Else If SameText(Trim(LPropName), 'Tipo') Then
          Begin
            If SameText(AValue, 'Code128') Then
              TLabelBarcode(LItem).Kind := bkCode128
            Else If SameText(AValue, 'EAN13') Then
              TLabelBarcode(LItem).Kind := bkEAN13
            Else If SameText(AValue, 'QR') Then
              TLabelBarcode(LItem).Kind := bkQR;
          End
          Else If SameText(Trim(LPropName), 'Mostrar Texto') Then
            TLabelBarcode(LItem).ShowText := StrToBoolDef(AValue, TLabelBarcode(LItem).ShowText)
          Else If SameText(Trim(LPropName), 'Módulo') Then
            TLabelBarcode(LItem).Module := StrToFloatDef(AValue, TLabelBarcode(LItem).Module)
          Else If SameText(Trim(LPropName), 'Ratio') Then
            TLabelBarcode(LItem).Ratio := StrToFloatDef(AValue, TLabelBarcode(LItem).Ratio);
        End
        Else If LItem IS TLabelImage Then
        Begin
          If SameText(Trim(LPropName), 'Ajustar') Then
            TLabelImage(LItem).Stretch := StrToBoolDef(AValue, TLabelImage(LItem).Stretch)
          Else If SameText(Trim(LPropName), 'Proporcional') Then
            TLabelImage(LItem).Proportional := StrToBoolDef(AValue, TLabelImage(LItem).Proportional);
        End
        Else If LItem IS TLabelGrid Then
        Begin
          LGrid := TLabelGrid(LItem);
          If LGrid.GetSelectionCount = 1 Then
          Begin
            LP := LGrid.GetFirstSelectedCell;
            LChildItem := LGrid.GetCellItem(LP.Y, LP.X);
            If Assigned(LChildItem) Then
            Begin
              // Aplicar ao objeto filho (propriedades sem prefixo)
              If (LSG.Cells[0, ARow] = 'Nome') Then
              Begin
                If LabelManager.ExistsName(AValue, LChildItem) Then
                Begin
                  ShowMessage('Já existe um objeto com o nome "' + AValue + '".');
                  Exit;
                End;
                LChildItem.Name := AValue;
              End
              Else If SameText(LPropName, 'Left (mm)') Then LChildItem.Left := MmToPx(StrToFloatDef(AValue, PxToMm(LChildItem.Left)))
              Else If SameText(LPropName, 'Top (mm)') Then LChildItem.Top := MmToPx(StrToFloatDef(AValue, PxToMm(LChildItem.Top)))
              Else If SameText(LPropName, 'Width (mm)') Then LChildItem.Width := MmToPx(StrToFloatDef(AValue, PxToMm(LChildItem.Width)))
              Else If SameText(LPropName, 'Height (mm)') Then LChildItem.Height := MmToPx(StrToFloatDef(AValue, PxToMm(LChildItem.Height)))
              Else If SameText(LPropName, 'Nome do campo') Then
              Begin
                If SameText(AValue, '<Vazio>') Then
                  LChildItem.FieldName := ''
                Else
                  LChildItem.FieldName := AValue;
              End
              Else If SameText(LPropName, 'Desvio de registos') Then LChildItem.RecordOffset := StrToIntDef(AValue, LChildItem.RecordOffset)
              Else If LChildItem IS TLabelText Then
              Begin
                If SameText(LPropName, 'Texto') Then TLabelText(LChildItem).Text := AValue
                Else If SameText(LPropName, 'Font') Then TLabelText(LChildItem).Font.Name := AValue // Fallback
                Else If SameText(LPropName, 'Nome') Then TLabelText(LChildItem).Font.Name := AValue
                Else If SameText(LPropName, 'Tamanho') Then TLabelText(LChildItem).Font.Size := StrToIntDef(AValue, TLabelText(LChildItem).Font.Size)
                Else If SameText(LPropName, 'Cor') Then TLabelText(LChildItem).Font.Color := StringToColorDef(AValue, TLabelText(LChildItem).Font.Color)

                // Styles
                Else If SameText(LPropName, 'Negrito') Then
                Begin
                  If StrToBoolDef(AValue, False) Then TLabelText(LChildItem).Font.Style := TLabelText(LChildItem).Font.Style + [fsBold]
                  Else
                    TLabelText(LChildItem).Font.Style := TLabelText(LChildItem).Font.Style - [fsBold];
                End
                Else If SameText(LPropName, 'Sublinhado') Then
                Begin
                  If StrToBoolDef(AValue, False) Then TLabelText(LChildItem).Font.Style := TLabelText(LChildItem).Font.Style + [fsUnderline]
                  Else
                    TLabelText(LChildItem).Font.Style := TLabelText(LChildItem).Font.Style - [fsUnderline];
                End
                Else If SameText(LPropName, 'Itálico') Then
                Begin
                  If StrToBoolDef(AValue, False) Then TLabelText(LChildItem).Font.Style := TLabelText(LChildItem).Font.Style + [fsItalic]
                  Else
                    TLabelText(LChildItem).Font.Style := TLabelText(LChildItem).Font.Style - [fsItalic];
                End
                Else If SameText(LPropName, 'Rasurado') Then
                Begin
                  If StrToBoolDef(AValue, False) Then TLabelText(LChildItem).Font.Style := TLabelText(LChildItem).Font.Style + [fsStrikeOut]
                  Else
                    TLabelText(LChildItem).Font.Style := TLabelText(LChildItem).Font.Style - [fsStrikeOut];
                End

                Else If SameText(LPropName, 'WordWrap') Then TLabelText(LChildItem).WordWrap := StrToBoolDef(AValue, TLabelText(LChildItem).WordWrap)
                Else If SameText(LPropName, 'Alinham. H.') Then
                Begin
                  If SameText(AValue, 'Esquerda') Then TLabelText(LChildItem).Alignment := taLeftJustify
                  Else If SameText(AValue, 'Centro') Then TLabelText(LChildItem).Alignment := taCenter
                  Else If SameText(AValue, 'Direita') Then TLabelText(LChildItem).Alignment := taRightJustify;
                End
                Else If SameText(LPropName, 'Alinham. V.') Then
                Begin
                  If SameText(AValue, 'Topo') Then TLabelText(LChildItem).VertAlignment := tlTop
                  Else If SameText(AValue, 'Centro') Then TLabelText(LChildItem).VertAlignment := tlCenter
                  Else If SameText(AValue, 'Base') Then TLabelText(LChildItem).VertAlignment := tlBottom;
                End;
              End
              Else If LChildItem IS TLabelBarcode Then
              Begin
                If SameText(LPropName, 'Texto') Then TLabelBarcode(LChildItem).Text := AValue
                Else If SameText(LPropName, 'Tipo') Then
                Begin
                  If SameText(AValue, 'Code128') Then TLabelBarcode(LChildItem).Kind := bkCode128
                  Else If SameText(AValue, 'EAN13') Then TLabelBarcode(LChildItem).Kind := bkEAN13
                  Else If SameText(AValue, 'QR') Then TLabelBarcode(LChildItem).Kind := bkQR;
                End
                Else If SameText(LPropName, 'Mostrar Texto') Then TLabelBarcode(LChildItem).ShowText := StrToBoolDef(AValue, TLabelBarcode(LChildItem).ShowText)
                Else If SameText(LPropName, 'Módulo') Then TLabelBarcode(LChildItem).Module := StrToFloatDef(AValue, TLabelBarcode(LChildItem).Module)
                Else If SameText(LPropName, 'Ratio') Then TLabelBarcode(LChildItem).Ratio := StrToFloatDef(AValue, TLabelBarcode(LChildItem).Ratio);
              End
              Else If LChildItem IS TLabelImage Then
              Begin
                If SameText(LPropName, 'Ajustar') Then
                  TLabelImage(LChildItem).Stretch := StrToBoolDef(AValue, TLabelImage(LChildItem).Stretch)
                Else If SameText(LPropName, 'Proporcional') Then
                  TLabelImage(LChildItem).Proportional := StrToBoolDef(AValue, TLabelImage(LChildItem).Proportional);
              End;
            End;
          End;

          // Propriedades da Célula (aplicar a todas as células selecionadas)
          If LGrid.GetSelectionCount > 0 Then
          Begin
            For Lr := 0 To LGrid.Rows - 1 Do
              For Lc := 0 To LGrid.Cols - 1 Do
                If LGrid.IsSelected(Lr, Lc) Then
                Begin
                  If SameText(LPropName, 'Cor Fundo') Then LGrid.SetCellBackColor(Lr, Lc, StringToColorDef(AValue, LGrid.GetCellBackColor(Lr, Lc)))
                  Else If SameText(LPropName, 'Cor Borda') Then LGrid.SetCellBorderColor(Lr, Lc, StringToColorDef(AValue, LGrid.GetCellBorderColor(Lr, Lc)))
                  Else If SameText(LPropName, 'Esquerda') Then LGrid.SetEdgeVisibility(Lr, Lc, beLeft, StrToBoolDef(AValue, True))
                  Else If SameText(LPropName, 'Topo') Then LGrid.SetEdgeVisibility(Lr, Lc, beTop, StrToBoolDef(AValue, True))
                  Else If SameText(LPropName, 'Direita') Then LGrid.SetEdgeVisibility(Lr, Lc, beRight, StrToBoolDef(AValue, True))
                  Else If SameText(LPropName, 'Baixo') Then LGrid.SetEdgeVisibility(Lr, Lc, beBottom, StrToBoolDef(AValue, True));
                End;
          End;
        End;
      End;

      UpdateObjectsFromData;
      GetActivePaintBox.Invalidate;
    Except
      SysUtils.Beep;
    End;
  Finally
    LSelectedList.Free;
  End;
End;


Procedure TfEtiqueta.RebuildObjectTree;
Var
  LRootNode: TTreeNode;
  Li: Integer;
  LTV: TTreeView;

  Procedure AddItemToTree(AParentNode: TTreeNode; AItem: TLabelItem);
  Var
    LNode: TTreeNode;
    Lr, Lc: Integer;
    LGrid: TLabelGrid;
    LSubItem: TLabelItem;
  Begin
    LNode := LTV.Items.AddChild(AParentNode, AItem.Name);
    LNode.Data := Pointer(AItem);

    If AItem IS TLabelGrid Then
    Begin
      LGrid := TLabelGrid(AItem);
      For Lr := 0 To LGrid.Rows - 1 Do
        For Lc := 0 To LGrid.Cols - 1 Do
        Begin
          LSubItem := LGrid.GetCellItem(Lr, Lc);
          If Assigned(LSubItem) Then
            AddItemToTree(LNode, LSubItem);
        End;
    End;
  End;

Begin
  LTV := GetActiveTreeView;
  If LTV = nil Then Exit;

  If LabelManager = nil Then Exit;

  FUpdatingTree := True;
  Try
    LTV.Items.BeginUpdate;
    LTV.Items.Clear;

    LRootNode := LTV.Items.Add(nil, 'Etiqueta');
    LRootNode.Data := nil;

    For Li := 0 To LabelManager.Count - 1 Do
      AddItemToTree(LRootNode, LabelManager[Li]);

    LRootNode.Expand(True);
  Finally
    LTV.Items.EndUpdate;
    FUpdatingTree := False;
  End;
End;

Procedure TfEtiqueta.SelectInTree(AItem: TLabelItem);
Var
  i: Integer;
  LNode: TTreeNode;
  LTV: TTreeView;
Begin
  LTV := GetActiveTreeView;
  If LTV = nil Then Exit;

  If FUpdatingTree Then Exit;
  FUpdatingTree := True;
  Try
    LNode := nil;

    // 1. Procurar por correspondência de ponteiro (Data)
    For i := 0 To LTV.Items.Count - 1 Do
    Begin
      If LTV.Items[i].Data = AItem Then
      Begin
        LNode := LTV.Items[i];
        Break;
      End;
    End;

    // 2. Salvaguarda: Procurar por nome se o ponteiro falhar (e AItem não for nil)
    If (LNode = nil) AND (AItem <> nil) Then
    Begin
      For i := 0 To LTV.Items.Count - 1 Do
      Begin
        If SameText(LTV.Items[i].Text, AItem.Name) Then
        Begin
          LNode := LTV.Items[i];
          // Atualizar o Data para futuras pesquisas
          LNode.Data := Pointer(AItem);
          Break;
        End;
      End;
    End;

    // 3. Se AItem for nil ou não encontrado, selecionar a raiz "Etiqueta"
    If (LNode = nil) AND (LTV.Items.Count > 0) Then
      LNode := LTV.Items[0];

    // Aplicar seleção
    If Assigned(LNode) Then
    Begin
      LTV.Selected := LNode;
      LNode.MakeVisible;
    End;
  Finally
    FUpdatingTree := False;
  End;
End;

Procedure TfEtiqueta.tvObjetosSelectionChanged(ASender: TObject);
Var
  LTV: TTreeView;
Begin
  LTV := ASender AS TTreeView;
  tvObjetosChange(ASender, LTV.Selected);
End;

Procedure TfEtiqueta.tvObjetosChange(ASender: TObject; ANode: TTreeNode);
Var
  LItem: TLabelItem;
  LTV: TTreeView;
Begin
  If FUpdatingTree Then Exit;
  If LabelManager = nil Then Exit;
  LTV := ASender AS TTreeView;

  // 1. Limpar TODAS as seleções atuais (objetos e células)
  LabelManager.DeselectAll;

  // 2. Identificar o novo item selecionado
  If (LTV.Selected = nil) OR (LTV.Selected.Data = nil) Then
  Begin
    // Selecionou a raiz "Etiqueta" ou nó inválido
    FSelectedItem := nil;
    ShowLabelProperties;
  End
  Else
  Begin
    // Selecionou um objeto real
    LItem := TLabelItem(LTV.Selected.Data);
    FSelectedItem := LItem;
    LItem.Selected := True;
    UpdatePropertiesPanel(LItem);
  End;

  // 3. Forçar redesenho do designer
  GetActivePaintBox.Invalidate;
End;

Procedure TfEtiqueta.DrawSelectionHandles(ACanvas: TCanvas; AItem: TLabelItem);
Const
  HandleSize = 6;
Var
  LR: TRect;
  LMidX, LMidY: Integer;
Begin
  If NOT Assigned(AItem) Then Exit;

  LR := Rect(AItem.Left, AItem.Top, AItem.Left + AItem.Width, AItem.Top + AItem.Height);
  LMidX := (LR.Left + LR.Right) DIV 2;
  LMidY := (LR.Top + LR.Bottom) DIV 2;

  ACanvas.Pen.Color := clBlue;
  ACanvas.Pen.Style := psSolid;
  ACanvas.Brush.Color := clWhite;

  // Draw 8 handles
  // Top-Left
  ACanvas.Rectangle(LR.Left - HandleSize DIV 2, LR.Top - HandleSize DIV 2,
    LR.Left + HandleSize DIV 2, LR.Top + HandleSize DIV 2);
  // Top
  ACanvas.Rectangle(LMidX - HandleSize DIV 2, LR.Top - HandleSize DIV 2,
    LMidX + HandleSize DIV 2, LR.Top + HandleSize DIV 2);
  // Top-Right
  ACanvas.Rectangle(LR.Right - HandleSize DIV 2, LR.Top - HandleSize DIV 2,
    LR.Right + HandleSize DIV 2, LR.Top + HandleSize DIV 2);
  // Right
  ACanvas.Rectangle(LR.Right - HandleSize DIV 2, LMidY - HandleSize DIV 2,
    LR.Right + HandleSize DIV 2, LMidY + HandleSize DIV 2);
  // Bottom-Right
  ACanvas.Rectangle(LR.Right - HandleSize DIV 2, LR.Bottom - HandleSize DIV 2,
    LR.Right + HandleSize DIV 2, LR.Bottom + HandleSize DIV 2);
  // Bottom
  ACanvas.Rectangle(LMidX - HandleSize DIV 2, LR.Bottom - HandleSize DIV 2,
    LMidX + HandleSize DIV 2, LR.Bottom + HandleSize DIV 2);
  // Bottom-Left
  ACanvas.Rectangle(LR.Left - HandleSize DIV 2, LR.Bottom - HandleSize DIV 2,
    LR.Left + HandleSize DIV 2, LR.Bottom + HandleSize DIV 2);
  // Left
  ACanvas.Rectangle(LR.Left - HandleSize DIV 2, LMidY - HandleSize DIV 2,
    LR.Left + HandleSize DIV 2, LMidY + HandleSize DIV 2);
End;

Function TfEtiqueta.DetectHandle(AItem: TLabelItem; AX, AY: Integer): TResizeHandle;
Const
  HandleSize = 6;
Var
  LR: TRect;
  LMidX, LMidY: Integer;

  Function InRect(AX, AY, ACX, ACY: Integer): Boolean;
  Begin
    Result := (Abs(AX - ACX) <= HandleSize) AND (Abs(AY - ACY) <= HandleSize);
  End;

Begin
  Result := rhNone;
  If NOT Assigned(AItem) Then Exit;

  LR := Rect(AItem.Left, AItem.Top, AItem.Left + AItem.Width, AItem.Top + AItem.Height);
  LMidX := (LR.Left + LR.Right) DIV 2;
  LMidY := (LR.Top + LR.Bottom) DIV 2;

  // Check handles in priority order (corners first for better UX)
  If InRect(AX, AY, LR.Left, LR.Top) Then Result := rhTopLeft
  Else If InRect(AX, AY, LR.Right, LR.Top) Then Result := rhTopRight
  Else If InRect(AX, AY, LR.Right, LR.Bottom) Then Result := rhBottomRight
  Else If InRect(AX, AY, LR.Left, LR.Bottom) Then Result := rhBottomLeft
  Else If InRect(AX, AY, LMidX, LR.Top) Then Result := rhTop
  Else If InRect(AX, AY, LR.Right, LMidY) Then Result := rhRight
  Else If InRect(AX, AY, LMidX, LR.Bottom) Then Result := rhBottom
  Else If InRect(AX, AY, LR.Left, LMidY) Then Result := rhLeft;
End;

Procedure TfEtiqueta.miDeleteClick(ASender: TObject);
Begin
  DeleteSelectedItem;
End;

Procedure TfEtiqueta.DeleteSelectedItem;
Var
  LSelectedList: TList;
  Li, Lr, Lc, Lj: Integer;
  LGrid: TLabelGrid;
  LItem: TLabelItem;
  LFound: Boolean;
Begin
  LSelectedList := LabelManager.GetSelectedItems;
  Try
    If LSelectedList.Count = 0 Then Exit;

    For Lj := 0 To LSelectedList.Count - 1 Do
    Begin
      LItem := TLabelItem(LSelectedList[Lj]);
      LFound := False;

      // Check if inside a grid
      For Li := 0 To LabelManager.Count - 1 Do
      Begin
        If LabelManager[Li] IS TLabelGrid Then
        Begin
          LGrid := TLabelGrid(LabelManager[Li]);
          For Lr := 0 To LGrid.Rows - 1 Do
          Begin
            For Lc := 0 To LGrid.Cols - 1 Do
            Begin
              If LGrid.GetCellItem(Lr, Lc) = LItem Then
              Begin
                LGrid.SetCellItem(Lr, Lc, nil); // This frees the item
                LFound := True;
                Break;
              End;
            End;
            If LFound Then Break;
          End;
        End;
        If LFound Then Break;
      End;

      If NOT LFound Then
      Begin
        LabelManager.RemoveItem(LItem);
      End;
    End;

    FSelectedItem := nil;
    ShowLabelProperties;
    RebuildObjectTree;
    GetActivePaintBox.Invalidate;
  Finally
    LSelectedList.Free;
  End;
End;

Procedure TfEtiqueta.CopyToClipboard(ASender: TObject);
Begin
  If Assigned(FSelectedItem) Then
  Begin
    If Assigned(FClipboard) Then FClipboard.Free;
    FClipboard := FSelectedItem.Clone;
  End;
End;

Procedure TfEtiqueta.CutToClipboard(ASender: TObject);
Begin
  If Assigned(FSelectedItem) Then
  Begin
    CopyToClipboard(ASender);
    DeleteSelectedItem;
  End;
End;

Procedure TfEtiqueta.PasteFromClipboard(ASender: TObject);
Var
  LNewItem: TLabelItem;
  LGrid: TLabelGrid;
  LSelPoint: TPoint;
Begin
  If Assigned(FClipboard) Then
  Begin
    LNewItem := FClipboard.Clone;
    LNewItem.Name := LabelManager.GenerateUniqueName(LNewItem.Name + '_Copy');

    // Check if pasting into a Grid Cell
    If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
    Begin
      LGrid := TLabelGrid(FSelectedItem);
      LSelPoint := LGrid.GetFirstSelectedCell;

      If (LSelPoint.X >= 0) AND (LSelPoint.Y >= 0) Then
      Begin
        If LGrid.GetCellItem(LSelPoint.Y, LSelPoint.X) = nil Then
        Begin
          LGrid.SetCellItem(LSelPoint.Y, LSelPoint.X, LNewItem);

          FSelectedItem := LNewItem;
          UpdatePropertiesPanel(LNewItem);
          RebuildObjectTree;
          SelectInTree(LNewItem);
          GetActivePaintBox.Invalidate;
          Exit;
        End;
      End;
    End;

    // Default: Add to LabelManager
    LNewItem.Left := LNewItem.Left + 10;
    LNewItem.Top := LNewItem.Top + 10;

    // Add to manager
    LabelManager.AddItem(LNewItem);
    FSelectedItem := LNewItem;
    RebuildObjectTree;
    SelectInTree(FSelectedItem);
    GetActivePaintBox.Invalidate;
  End;
End;

Procedure TfEtiqueta.miDuplicateClick(ASender: TObject);
Var
  LNewItem: TLabelItem;
Begin
  If Assigned(FSelectedItem) Then
  Begin
    LNewItem := LabelManager.DuplicateItem(FSelectedItem);

    // Deselect all and select the new item
    LabelManager.DeselectAll;

    If Assigned(LNewItem) Then
    Begin
      LNewItem.Selected := True;
      FSelectedItem := LNewItem;
      RebuildObjectTree;
      SelectInTree(LNewItem);
      UpdatePropertiesPanel(LNewItem);
    End;

    GetActivePaintBox.Invalidate;
  End;
End;

Procedure TfEtiqueta.miBringToFrontClick(ASender: TObject);
Begin
  If Assigned(FSelectedItem) Then
  Begin
    LabelManager.BringToFront(FSelectedItem);
    GetActivePaintBox.Invalidate;
  End;
End;

Procedure TfEtiqueta.miSendToBackClick(ASender: TObject);
Begin
  If Assigned(FSelectedItem) Then
  Begin
    LabelManager.SendToBack(FSelectedItem);
    GetActivePaintBox.Invalidate;
  End;
End;

Procedure TfEtiqueta.miPropertiesClick(ASender: TObject);
Begin
  // Properties panel is already visible, just ensure it's updated
  If Assigned(FSelectedItem) Then
    UpdatePropertiesPanel(FSelectedItem);
End;

Procedure TfEtiqueta.miInsertColLeftClick(ASender: TObject);
Var
  LGrid: TLabelGrid;
  LP: TPoint;
Begin
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    LP := LGrid.GetFirstSelectedCell;
    If (LP.X >= 0) Then
    Begin
      LGrid.InsertCol(LP.X);
      GetActivePaintBox.Invalidate;
      UpdatePropertiesPanel(LGrid);
      RebuildObjectTree;
    End;
  End;
End;

Procedure TfEtiqueta.miInsertColRightClick(ASender: TObject);
Var
  LGrid: TLabelGrid;
  LP: TPoint;
Begin
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    LP := LGrid.GetFirstSelectedCell;
    If (LP.X >= 0) Then
    Begin
      LGrid.InsertCol(LP.X + 1);
      GetActivePaintBox.Invalidate;
      UpdatePropertiesPanel(LGrid);
      RebuildObjectTree;
    End;
  End;
End;

Procedure TfEtiqueta.miInsertRowAboveClick(ASender: TObject);
Var
  LGrid: TLabelGrid;
  LP: TPoint;
Begin
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    LP := LGrid.GetFirstSelectedCell;
    If (LP.Y >= 0) Then
    Begin
      LGrid.InsertRow(LP.Y);
      GetActivePaintBox.Invalidate;
      UpdatePropertiesPanel(LGrid);
      RebuildObjectTree;
    End;
  End;
End;

Procedure TfEtiqueta.miInsertRowBelowClick(ASender: TObject);
Var
  LGrid: TLabelGrid;
  LP: TPoint;
Begin
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    LP := LGrid.GetFirstSelectedCell;
    If (LP.Y >= 0) Then
    Begin
      LGrid.InsertRow(LP.Y + 1);
      GetActivePaintBox.Invalidate;
      UpdatePropertiesPanel(LGrid);
      RebuildObjectTree;
    End;
  End;
End;

Procedure TfEtiqueta.miDeleteColClick(ASender: TObject);
Var
  LGrid: TLabelGrid;
  LP: TPoint;
Begin
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    LP := LGrid.GetFirstSelectedCell;
    If (LP.X >= 0) Then
    Begin
      LGrid.DeleteCol(LP.X);
      GetActivePaintBox.Invalidate;
      UpdatePropertiesPanel(LGrid);
      RebuildObjectTree;
    End;
  End;
End;

Procedure TfEtiqueta.miDeleteRowClick(ASender: TObject);
Var
  LGrid: TLabelGrid;
  LP: TPoint;
Begin
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    LP := LGrid.GetFirstSelectedCell;
    If (LP.Y >= 0) Then
    Begin
      LGrid.DeleteRow(LP.Y);
      GetActivePaintBox.Invalidate;
      UpdatePropertiesPanel(LGrid);
      RebuildObjectTree;
    End;
  End;
End;

Procedure TfEtiqueta.pmContextPopup(ASender: TObject);
Var
  LGrid: TLabelGrid;
  Lr, Lc, Li: Integer;
  LItem: TLabelItem;
  LHasGrid: Boolean;
Begin
  If Assigned(miGroup) Then miGroup.Visible := False;
  If Assigned(miUngroup) Then miUngroup.Visible := False;
  If Assigned(miInsert) Then miInsert.Visible := False;
  If Assigned(miGridDelete) Then miGridDelete.Visible := False;

  LHasGrid := False;
  LGrid := nil;
  Lr := -1;
  Lc := -1;

  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    LHasGrid := True;
    If LGrid.GetSelectionCount > 0 Then
    Begin
      // Determine if we can group/ungroup
      If LGrid.CanGroupSelected Then miGroup.Visible := True;
      If LGrid.CanUngroupSelected Then miUngroup.Visible := True;

      // Determine if we can insert/delete (only for single cell selection or focus)
      // Actually, we can support it for any selection, taking the first cell as reference
      miInsert.Visible := True;
      miGridDelete.Visible := True;
    End;
  End;

  // Case 1: Grid itself is selected
  If (FSelectedItem <> nil) AND (FSelectedItem IS TLabelGrid) Then
  Begin
    LGrid := TLabelGrid(FSelectedItem);
    Lr := LGrid.GetFirstSelectedCell.Y;
    Lc := LGrid.GetFirstSelectedCell.X;
    LHasGrid := (Lr >= 0) AND (Lc >= 0);

    // New Grouping Logic
    If Assigned(miGroup) Then miGroup.Visible := LGrid.CanGroupSelected;
    If Assigned(miUngroup) Then miUngroup.Visible := LGrid.CanUngroupSelected;
  End
  Else If (FSelectedItem <> nil) Then
  Begin
    // Case 2: A child inside a grid is selected; find its parent grid and cell
    For Li := 0 To LabelManager.Count - 1 Do
    Begin
      LItem := LabelManager[Li];
      If LItem IS TLabelGrid Then
      Begin
        LGrid := TLabelGrid(LItem);
        For Lr := 0 To LGrid.Rows - 1 Do
        Begin
          For Lc := 0 To LGrid.Cols - 1 Do
          Begin
            If LGrid.GetCellItem(Lr, Lc) = FSelectedItem Then
            Begin
              LHasGrid := True;
              Break;
            End;
          End;
          If LHasGrid Then Break;
        End;
        If LHasGrid Then Break;
      End;
    End;
  End;
End;

Procedure TfEtiqueta.miDeleteTabClick(ASender: TObject);
Begin
  DeleteActiveTab;
End;

Procedure TfEtiqueta.DeleteActiveTab;
Var
  LTab: TTabSheet;
  LSettings: TTabSettings;
Begin
  LTab := pcEtiquetas.ActivePage;
  If (LTab = nil) OR (LTab = tsMais) Then Exit;

  If MessageDlg('Confirmar', 'Deseja eliminar esta etiqueta?', mtConfirmation, [mbYes, mbNo], 0) = mrYes Then
  Begin
    LSettings := TTabSettings(LTab.Tag);
    LTab.Free;
    If LSettings <> nil Then LSettings.Free;
    
    // Select previous tab if possible
    If pcEtiquetas.PageCount > 1 Then
      pcEtiquetas.ActivePageIndex := pcEtiquetas.PageCount - 2;
      
    ShowLabelProperties;
    RebuildObjectTree;
  End;
End;

Procedure TfEtiqueta.miGroupClick(ASender: TObject);
Begin
  If (FSelectedItem IS TLabelGrid) Then
  Begin
    TLabelGrid(FSelectedItem).GroupSelected;
    GetActivePaintBox.Invalidate;
  End;
End;

Procedure TfEtiqueta.miUngroupClick(ASender: TObject);
Begin
  If (FSelectedItem IS TLabelGrid) Then
  Begin
    TLabelGrid(FSelectedItem).UngroupSelected;
    GetActivePaintBox.Invalidate;
  End;
End;


{ TDataLoaderThread }

Constructor TDataLoaderThread.Create(AOwner: TForm; ADados: TDadosExternos; AGrid: TStringGrid; AFilters: TStringList; ACheckedKeys: TStringList; APerformAutoSize: Boolean);
Begin
  Inherited Create(False);
  FreeOnTerminate := False;
  FOwner := AOwner;
  FDados := ADados;
  FGrid := AGrid;
  FHeaders := TStringList.Create;
  FBatchRows := TList.Create;
  FIsFinished := False;
  FPerformAutoSize := APerformAutoSize;
  FFilters := TStringList.Create;
  If AFilters <> nil Then
    FFilters.Assign(AFilters);
  FCheckedKeys := TStringList.Create;
  If ACheckedKeys <> nil Then
    FCheckedKeys.Assign(ACheckedKeys);
End;

Destructor TDataLoaderThread.Destroy;
Var
  Li: Integer;
Begin
  FHeaders.Free;
  FFilters.Free;
  FCheckedKeys.Free;
  For Li := 0 To FBatchRows.Count - 1 Do
    TObject(FBatchRows[Li]).Free;
  FBatchRows.Free;
  Inherited Destroy;
End;

Procedure TDataLoaderThread.SyncBatch;
Var
  Lr, Lc: Integer;
  LRowData: TStringList;
  LfE: TfEtiqueta;
  LOldFilters: TStringList;
  LSelStart, LSelLen: Integer;
  LEditor: TCustomEdit;
  LRowKey: String;
  LCurrentGridRow: Integer;
Begin
  If Terminated OR NOT Assigned(FGrid) Then Exit;

  LfE := TfEtiqueta(FOwner);

  // Preservar posição do cursor se estivermos a editar a linha de filtros
  LSelStart := -1;
  LEditor := nil;
  If (FGrid.Row = 1) AND Assigned(FGrid.Editor) AND (FGrid.Editor IS TCustomEdit) Then
  Begin
    LEditor := TCustomEdit(FGrid.Editor);
    LSelStart := LEditor.SelStart;
    LSelLen := LEditor.SelLength;
  End;

  FGrid.BeginUpdate;
  Try
    If FHeaders.Count > 0 Then
    Begin
      If FGrid.ColCount <> FHeaders.Count + 1 Then
      Begin
        // Preservar os filtros existentes na linha 1 antes de alterar ColCount
        LOldFilters := TStringList.Create;
        Try
          If FGrid.RowCount >= 2 Then
          Begin
            For Lc := 0 To FGrid.ColCount - 1 Do
              LOldFilters.Add(FGrid.Cells[Lc, 1]);
          End;

          FGrid.ColCount := FHeaders.Count + 1;

          // Cabeçalho da coluna de Checkbox
          FGrid.Cells[0, 0] := '✅';
          FGrid.Cells[0, 1] := '0';
          FGrid.ColWidths[0] := iCheckColWidth;

          // Restaurar cabeçalhos de dados
          For Lc := 0 To FHeaders.Count - 1 Do
            FGrid.Cells[Lc + 1, 0] := FHeaders[Lc];

          // Restaurar filtros preservados
          If FGrid.RowCount >= 2 Then
          Begin
            For Lc := 0 To LOldFilters.Count - 1 Do
            Begin
              If Lc < FGrid.ColCount Then
                FGrid.Cells[Lc, 1] := LOldFilters[Lc];
            End;
          End;
        Finally
          LOldFilters.Free;
        End;
      End;
    End;

    For Lr := 0 To FBatchRows.Count - 1 Do
    Begin
      LRowData := TStringList(FBatchRows[Lr]);

      // Gerar chave para esta nova linha para ver se estava marcada
      LRowKey := '';
      For Lc := 0 To LRowData.Count - 1 Do
        LRowKey := LRowKey + LRowData[Lc] + '|';

      // Inserir '1' se estava marcada, '0' caso contrário
      If FCheckedKeys.IndexOf(LRowKey) >= 0 Then
        LRowData.Insert(0, '1')
      Else
        LRowData.Insert(0, '0');

      // Se o filtro de "Apenas Marcados" estiver ativo, saltar linhas desmarcadas
      If (FFilters.Values['_CHECKED_'] = '1') AND (LRowData[0] = '0') Then
      Begin
        LRowData.Free;
        Continue;
      End;

      LCurrentGridRow := FGrid.RowCount;
      FGrid.RowCount := LCurrentGridRow + 1;
      FGrid.Rows[LCurrentGridRow].Assign(LRowData);
      LRowData.Free;
    End;
    FBatchRows.Clear;

    If FIsFinished Then
    Begin
      // AutoSizeColumns apenas na primeira carga ou se solicitado
      If FPerformAutoSize Then
      Begin
        FGrid.AutoSizeColumns;
        FGrid.ColWidths[0] := 20;
        FGrid.AjustarLarguraAoLimite;
        TfEtiqueta(FOwner).FDataLoadedOnce := True;
      End;

      LfE.lDataStaus.Caption := IntToStr(FGrid.RowCount - 2) + ' registos';
      LfE.UpdateObjectsFromData;
    End
    Else
      LfE.lDataStaus.Caption := 'A carregar ' + IntToStr(FGrid.RowCount - 2) + ' registos...';

  Finally
    FGrid.EndUpdate;

    // Restaurar posição do cursor
    If (LSelStart >= 0) AND Assigned(LEditor) Then
    Begin
      LEditor.SelStart := LSelStart;
      LEditor.SelLength := LSelLen;
    End;
  End;
End;

Procedure TDataLoaderThread.SyncError;
Begin
  TfEtiqueta(FOwner).lDataStaus.Caption := 'ERRO: ' + FError;
  ShowMessage(FError);
End;

Procedure TDataLoaderThread.Execute;
Var
  LWorkbook: TsWorkbook;
  LWorksheet: TsWorksheet;
  LSourceStream, LDestStream: TFileStream;
  LTempFileName: String;
  Lr, Lc, Li, LColIdx: Integer;
  LConn: TSQLConnection;
  LQry: TSQLQuery;
  LTrans: TSQLTransaction;
  LRow: TStringList;
  LFullSQL, LFilterVal, LVal, LWhere: String;
  LPass: Boolean;
Begin
  Try
    Case FDados.TipoDadosExternos Of
      tdeFolhaCalculo:
      Begin
        If FileExists(FDados.FicheiroDados) Then
        Begin
          LWorkbook := TsWorkbook.Create;
          Try
            LTempFileName := IncludeTrailingPathDelimiter(GetTempDir) + 'LabelIt_Thread_' + ExtractFileName(FDados.FicheiroDados);
            LSourceStream := TFileStream.Create(FDados.FicheiroDados, fmOpenRead OR fmShareDenyNone);
            Try
              LDestStream := TFileStream.Create(LTempFileName, fmCreate);
              Try
                LDestStream.CopyFrom(LSourceStream, LSourceStream.Size);
              Finally
                LDestStream.Free;
              End;
            Finally
              LSourceStream.Free;
            End;

            LWorkbook.ReadFromFile(LTempFileName);
            If FileExists(LTempFileName) Then DeleteFile(LTempFileName);

            If (Length(FDados.Params) > 0) Then
              LWorksheet := LWorkbook.GetWorksheetByName(FDados.Params[0])
            Else
              LWorksheet := LWorkbook.GetWorksheetByIndex(0);

            If Assigned(LWorksheet) Then
            Begin
              For Lc := 0 To LWorksheet.GetLastColIndex Do
                FHeaders.Add(LWorksheet.ReadAsText(LWorksheet.GetCell(0, Lc)));

              For Lr := 1 To LWorksheet.GetLastRowIndex Do
              Begin
                // Aplicar filtros se existirem
                LPass := True;
                If (FFilters <> nil) AND (FFilters.Count > 0) Then
                Begin
                  For Li := 0 To FFilters.Count - 1 Do
                  Begin
                    If Pos('_', FFilters.Names[Li]) = 1 Then Continue; // Ignorar filtros locais (ex: _CHECKED_)
                    LColIdx := FHeaders.IndexOf(FFilters.Names[Li]);
                    If LColIdx >= 0 Then
                    Begin
                      LVal := LWorksheet.ReadAsText(LWorksheet.GetCell(Lr, LColIdx));
                      // Pesquisa parcial case-insensitive
                      If Pos(UpperCase(FFilters.ValueFromIndex[Li]), UpperCase(LVal)) = 0 Then
                      Begin
                        LPass := False;
                        Break;
                      End;
                    End;
                  End;
                End;

                If NOT LPass Then Continue;

                LRow := TStringList.Create;
                For Lc := 0 To LWorksheet.GetLastColIndex Do
                  LRow.Add(LWorksheet.ReadAsText(LWorksheet.GetCell(Lr, Lc)));
                FBatchRows.Add(LRow);

                If (FBatchRows.Count >= 5000) OR (Lr = LWorksheet.GetLastRowIndex) Then
                  Synchronize(SyncBatch);

                If Terminated Then Break;
              End;
            End;
          Finally
            LWorkbook.Free;
          End;
        End;
      End;

      tdeSQLite, tdeSQLServer, tdeMySQL:
      Begin
        LConn := nil;
        Try
          Case FDados.TipoDadosExternos Of
            tdeSQLite:
            Begin
              LConn := TSQLite3Connection.Create(nil);
              LConn.DatabaseName := FDados.FicheiroDados;
            End;
            tdeSQLServer:
            Begin
              LConn := TMSSQLConnection.Create(nil);
              LConn.HostName := FDados.Servidor;
              LConn.DatabaseName := FDados.BaseDeDados;
              LConn.UserName := FDados.Utilizador;
              LConn.Password := FDados.PalavraPasse;
            End;
            tdeMySQL:
            Begin
              Case FDados.Versao Of
                msv40: LConn := TMySQL40Connection.Create(nil);
                msv41: LConn := TMySQL41Connection.Create(nil);
                msv50: LConn := TMySQL50Connection.Create(nil);
                msv51: LConn := TMySQL51Connection.Create(nil);
                msv55: LConn := TMySQL55Connection.Create(nil);
                msv56: LConn := TMySQL56Connection.Create(nil);
                msv57: LConn := TMySQL57Connection.Create(nil);
                msv80: LConn := TMySQL80Connection.Create(nil);
              End;
              If Assigned(LConn) Then
              Begin
                LConn.HostName := FDados.Servidor;
                LConn.DatabaseName := FDados.BaseDeDados;
                LConn.UserName := FDados.Utilizador;
                LConn.Password := FDados.PalavraPasse;
              End;
            End;
          End;

          If Assigned(LConn) Then
          Begin
            Try
              LTrans := TSQLTransaction.Create(nil);
              LTrans.Database := LConn;
              LQry := TSQLQuery.Create(nil);
              LQry.Database := LConn;
              LQry.Transaction := LTrans;
              Try
                LFullSQL := '';
                For Li := 0 To High(FDados.Query) Do
                  LFullSQL := LFullSQL + FDados.Query[Li] + ' ';

                LFullSQL := Trim(LFullSQL);
                If (LFullSQL <> '') AND (LFullSQL[Length(LFullSQL)] = ';') Then
                  Delete(LFullSQL, Length(LFullSQL), 1);

                // Se houver filtros, envolvemos a query original numa subquery
                LWhere := '';
                If (FFilters <> nil) AND (FFilters.Count > 0) Then
                Begin
                  For Li := 0 To FFilters.Count - 1 Do
                  Begin
                    If Pos('_', FFilters.Names[Li]) = 1 Then Continue; // Ignorar filtros locais (ex: _CHECKED_)

                    If LWhere <> '' Then LWhere := LWhere + ' AND ';

                    // Pesquisa parcial case-insensitive com wildcards automáticos
                    LFilterVal := UpperCase(FFilters.ValueFromIndex[Li]);
                    If Pos('%', LFilterVal) = 0 Then
                      LFilterVal := '%' + LFilterVal + '%';

                    LWhere := LWhere + 'UPPER("' + FFilters.Names[Li] + '") LIKE ' + QuotedStr(LFilterVal);
                  End;
                End;

                If LWhere <> '' Then
                Begin
                  If (FDados.TipoDadosExternos = tdeSQLServer) AND (Pos('ORDER BY', UpperCase(LFullSQL)) > 0) Then
                    LFullSQL := 'SELECT * FROM (SELECT TOP 100 PERCENT ' + Copy(LFullSQL, 8, Length(LFullSQL)) + ') AS SubQ WHERE ' + LWhere
                  Else
                    LFullSQL := 'SELECT * FROM (' + LFullSQL + ') AS SubQ WHERE ' + LWhere;
                End;

                LQry.SQL.Text := LFullSQL;
                LQry.Open;

                For Li := 0 To LQry.FieldCount - 1 Do
                  FHeaders.Add(LQry.Fields[Li].FieldName);

                While NOT LQry.EOF Do
                Begin
                  LRow := TStringList.Create;
                  For Li := 0 To LQry.FieldCount - 1 Do
                    LRow.Add(LQry.Fields[Li].AsString);
                  FBatchRows.Add(LRow);

                  If (FBatchRows.Count >= 5000) Then
                    Synchronize(SyncBatch);

                  If Terminated Then Break;
                  LQry.Next;
                End;
              Finally
                LQry.Free;
                LTrans.Free;
              End;
            Finally
              LConn.Free;
            End;
          End;
        Except
          On E: Exception Do
          Begin
            FError := 'Erro SQL: ' + E.Message;
            Synchronize(SyncError);
          End;
        End;
      End;
    End;

    FIsFinished := True;
    Synchronize(SyncBatch);

  Except
    On E: Exception Do
    Begin
      FError := 'Erro ao carregar dados: ' + E.Message;
      Synchronize(SyncError);
    End;
  End;
End;

End.

