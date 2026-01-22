Unit uLabelObjects;

{$mode Delphi}

Interface

Uses
  //LabelIt
  ubarcodes,
  //Lazarus
  Classes, SysUtils, Windows, Graphics, Contnrs, Menus, Types;

Type
  TLabelItemType = (litText, litImage, litBarcode, litGrid);

  TBarcodeKind = (
    bkCode128,
    bkEAN13,
    bkQR
    );

  // Border visibility flags for each edge
  TBorderEdge = (beLeft, beTop, beRight, beBottom);
  TBorderEdges = Set Of TBorderEdge;

  // Cell formatting structure
  TCellFormat = Record
    BackColor: TColor;          // Background color
    BorderColor: TColor;        // Border color
    BorderVisible: TBorderEdges; // Which borders are visible
  End;

  // Merged cell information
  TMergedCell = Record
    Row: Integer;      // Starting row
    Col: Integer;      // Starting column
    RowSpan: Integer;  // Number of rows to span
    ColSpan: Integer;  // Number of columns to span
  End;

  TLabelGrid = Class;

  TLabelItem = Class
  Private
    FName: String;
    FLeft: Integer;
    FTop: Integer;
    FWidth: Integer;
    FHeight: Integer;
    FSelected: Boolean;
    FVisible: Boolean;
    FPopup: TPopupMenu;
    FFieldName: String;
    FRecordOffset: Integer;

    Procedure SetLeft(AValue: Integer); Virtual;
    Procedure SetTop(AValue: Integer); Virtual;
    Procedure SetWidth(AValue: Integer); Virtual;
    Procedure SetHeight(AValue: Integer); Virtual;
  Public
    Constructor Create; Virtual;
    Procedure Draw(ACanvas: TCanvas); Virtual; Abstract;
    Function HitTest(AX, AY: Integer): Boolean; Virtual;
    Function Clone: TLabelItem; Virtual; Abstract;

    Procedure SaveToWriter(AWriter: TWriter); Virtual;
    Procedure LoadFromReader(AReader: TReader); Virtual;
    Function GetItemType: TLabelItemType; Virtual; Abstract;

    Class Function CreateFromReader(AReader: TReader): TLabelItem; Static;

    Property Name: String Read FName Write FName;
    Property Left: Integer Read FLeft Write SetLeft;
    Property Top: Integer Read FTop Write SetTop;
    Property Width: Integer Read FWidth Write SetWidth;
    Property Height: Integer Read FHeight Write SetHeight;
    Property Selected: Boolean Read FSelected Write FSelected;
    Property Visible: Boolean Read FVisible Write FVisible;
    Property Popup: TPopupMenu Read FPopup Write FPopup;
    Property FieldName: String Read FFieldName Write FFieldName;
    Property RecordOffset: Integer Read FRecordOffset Write FRecordOffset;
  End;

  TLabelText = Class(TLabelItem)
  Private
    FText: String;
    FFont: TFont;
    FAlignment: TAlignment;
    FVertAlignment: TTextLayout;
    FWordWrap: Boolean;
    Procedure SetText(AValue: String);
  Public
    Constructor Create; Override;
    Destructor Destroy; Override;
    Procedure Draw(ACanvas: TCanvas); Override;
    Function Clone: TLabelItem; Override;
    Procedure SaveToWriter(AWriter: TWriter); Override;
    Procedure LoadFromReader(AReader: TReader); Override;
    Function GetItemType: TLabelItemType; Override;

    Property Text: String Read FText Write SetText;
    Property Font: TFont Read FFont;
    Property Alignment: TAlignment Read FAlignment Write FAlignment;
    Property VertAlignment: TTextLayout Read FVertAlignment Write FVertAlignment;
    Property WordWrap: Boolean Read FWordWrap Write FWordWrap;
  End;

  TLabelImage = Class(TLabelItem)
  Private
    FPicture: TPicture;
    FStretch: Boolean;
    FProportional: Boolean;
  Public
    Constructor Create; Override;
    Destructor Destroy; Override;
    Procedure Draw(ACanvas: TCanvas); Override;
    Function Clone: TLabelItem; Override;
    Procedure SaveToWriter(AWriter: TWriter); Override;
    Procedure LoadFromReader(AReader: TReader); Override;
    Function GetItemType: TLabelItemType; Override;

    Property Picture: TPicture Read FPicture Write FPicture;
    Property Stretch: Boolean Read FStretch Write FStretch;
    Property Proportional: Boolean Read FProportional Write FProportional;
  End;

  TLabelBarcode = Class(TLabelItem)
  Private
    FText: String;
    FShowText: Boolean;
    FKind: TBarcodeKind;
    FModule: Double;
    FRatio: Double;

    Procedure UpdateDimensions;
    Procedure SetText(AValue: String);
    Procedure SetKind(AValue: TBarcodeKind);
    Procedure SetModule(AValue: Double);
    Procedure SetRatio(AValue: Double);
    Procedure SetShowText(AValue: Boolean);
  Protected
    Procedure SetWidth(AValue: Integer); Override;
  Public
    Constructor Create; Override;
    Procedure Draw(ACanvas: TCanvas); Override;
    Function Clone: TLabelItem; Override;
    Procedure SaveToWriter(AWriter: TWriter); Override;
    Procedure LoadFromReader(AReader: TReader); Override;
    Function GetItemType: TLabelItemType; Override;

    Property Text: String Read FText Write SetText;
    Property ShowText: Boolean Read FShowText Write SetShowText;
    Property Kind: TBarcodeKind Read FKind Write SetKind;
    Property Module: Double Read FModule Write SetModule;
    Property Ratio: Double Read FRatio Write SetRatio;
  End;

  TLabelGrid = Class(TLabelItem)
  Private
    FRows: Integer;
    FCols: Integer;
    FCells: Array Of Array Of TLabelItem;
    FCellFormats: Array Of Array Of TCellFormat;  // Cell formatting
    FMergedCells: Array Of TMergedCell;           // Merged cell regions
    FSelectedCells: Array Of TPoint;              // Multiple selected cells
    FColWidths: Array Of Integer;                 // Individual column widths
    FRowHeights: Array Of Integer;                // Individual row heights

    Procedure SetRows(AValue: Integer);
    Procedure SetCols(AValue: Integer);
    Procedure UpdateCells;
    Procedure UpdateFormats;
    Procedure NormalizeMerges;
    Procedure InitializeCellFormats;
    Function GetActualCell(ARow, ACol: Integer; Var AActualRow, AActualCol: Integer): Boolean;
    Procedure UpdateDimensions;
  Protected
    Procedure SetWidth(AValue: Integer); Override;
    Procedure SetHeight(AValue: Integer); Override;
  Public
    Constructor Create; Override;
    Destructor Destroy; Override;
    Procedure Draw(ACanvas: TCanvas); Override;
    Function Clone: TLabelItem; Override;
    Procedure SaveToWriter(AWriter: TWriter); Override;
    Procedure LoadFromReader(AReader: TReader); Override;
    Function GetItemType: TLabelItemType; Override;

    Function IsCellMerged(ARow, ACol: Integer; Var AMergeInfo: TMergedCell): Boolean;

    Property Rows: Integer Read FRows Write SetRows;
    Property Cols: Integer Read FCols Write SetCols;

    Procedure SetCellItem(ARow, ACol: Integer; AItem: TLabelItem);
    Function GetCellItem(ARow, ACol: Integer): TLabelItem;
    Function FindChildAt(AX, AY: Integer): TLabelItem;
    Function GetCellAt(AX, AY: Integer; Out ARow, ACol: Integer): Boolean;

    // Selection methods
    Procedure ClearSelection;
    Procedure SelectCell(ARow, ACol: Integer; AAdd: Boolean = False);
    Procedure ToggleCellSelection(ARow, ACol: Integer);
    Procedure SelectRange(ARow1, ACol1, ARow2, ACol2: Integer; AAdd: Boolean = False);
    Function IsSelected(ARow, ACol: Integer): Boolean;
    Function GetSelectionCount: Integer;
    Function GetFirstSelectedCell: TPoint;

    // Grouping/Ungrouping checks and actions
    Function CanGroupSelected: Boolean;
    Function CanUngroupSelected: Boolean;
    Procedure GroupSelected;
    Procedure UngroupSelected;

    // Cell formatting methods
    Procedure SetCellBackColor(ARow, ACol: Integer; AColor: TColor);
    Function GetCellBackColor(ARow, ACol: Integer): TColor;
    Procedure SetCellBorderColor(ARow, ACol: Integer; AColor: TColor);
    Function GetCellBorderColor(ARow, ACol: Integer): TColor;
    Procedure SetCellBorderVisible(ARow, ACol: Integer; AEdges: TBorderEdges);
    Procedure SetEdgeVisibility(ARow, ACol: Integer; AEdge: TBorderEdge; AVisible: Boolean);
    Function GetCellBorderVisible(ARow, ACol: Integer): TBorderEdges;

    // Merge operations
    Procedure MergeCells(AStartRow, AStartCol, ARowSpan, AColSpan: Integer);
    Procedure UnmergeCells(ARow, ACol: Integer);
    Procedure ClearAllMerges;

    // Structure modification
    Procedure InsertRow(AIndex: Integer);
    Procedure InsertCol(AIndex: Integer);
    Procedure DeleteRow(AIndex: Integer);
    Procedure DeleteCol(AIndex: Integer);

    // Grid line resizing support
    Function HitTestGridLine(AX, AY: Integer; Var AIsRow: Boolean; Var AIndex: Integer): Boolean;
    Procedure SetColWidth(AIndex, AWidth: Integer);
    Procedure SetRowHeight(AIndex, AHeight: Integer);
    Function GetColWidth(AIndex: Integer): Integer;
    Function GetRowHeight(AIndex: Integer): Integer;
  End;

  TLabelManager = Class
  Private
    FItems: TObjectList;
    Function GetItem(AIndex: Integer): TLabelItem;
    Function GetCount: Integer;
  Public
    Constructor Create;
    Destructor Destroy; Override;
    Procedure Clear;
    Procedure AddItem(AItem: TLabelItem);
    Function AddText(AText: String): TLabelText;
    Function AddImage: TLabelImage;
    Function AddBarcode(AText: String; AKind: TBarcodeKind): TLabelBarcode;
    Function AddGrid(ARows, ACols: Integer): TLabelGrid;
    Procedure RemoveItem(AItem: TLabelItem);
    Function ExistsName(Const AName: String; AExcludeItem: TLabelItem = nil): Boolean;
    Function GenerateUniqueName(Const APrefix: String): String;
    Function ExtractItem(AItem: TLabelItem): TLabelItem;
    Procedure DrawAll(ACanvas: TCanvas);
    Function FindItemAt(AX, AY: Integer): TLabelItem;
    Function FindGridAt(AX, AY: Integer; AExclude: TLabelItem): TLabelGrid;
    Procedure BringToFront(AItem: TLabelItem);
    Procedure SendToBack(AItem: TLabelItem);
    Function DuplicateItem(AItem: TLabelItem): TLabelItem;
    Function GetSelectedItems: TList;
    Procedure DeselectAll;

    Property Items[AIndex: Integer]: TLabelItem Read GetItem; Default;
    Property Count: Integer Read GetCount;
  End;

Implementation

{ TLabelItem }

Constructor TLabelItem.Create;
Begin
  FVisible := True;
  FSelected := False;
  FLeft := 10;
  FTop := 10;
  FWidth := 100;
  FHeight := 50;
  FFieldName := '';
  FRecordOffset := 0;
End;

Procedure TLabelItem.SetLeft(AValue: Integer);
Begin
  FLeft := AValue;
End;

Procedure TLabelItem.SetTop(AValue: Integer);
Begin
  FTop := AValue;
End;

Procedure TLabelItem.SetWidth(AValue: Integer);
Begin
  FWidth := AValue;
End;

Procedure TLabelItem.SetHeight(AValue: Integer);
Begin
  FHeight := AValue;
End;

Function TLabelItem.HitTest(AX, AY: Integer): Boolean;
Begin
  Result := (AX >= FLeft) AND (AX <= FLeft + FWidth) AND (AY >= FTop) AND (AY <= FTop + FHeight);
End;

Procedure TLabelItem.SaveToWriter(AWriter: TWriter);
Begin
  AWriter.WriteString(FName);
  AWriter.WriteInteger(FLeft);
  AWriter.WriteInteger(FTop);
  AWriter.WriteInteger(FWidth);
  AWriter.WriteInteger(FHeight);
  AWriter.WriteBoolean(FVisible);
  AWriter.WriteString(FFieldName);
  AWriter.WriteInteger(FRecordOffset);
End;

Procedure TLabelItem.LoadFromReader(AReader: TReader);
Begin
  FName := AReader.ReadString;
  FLeft := AReader.ReadInteger;
  FTop := AReader.ReadInteger;
  FWidth := AReader.ReadInteger;
  FHeight := AReader.ReadInteger;
  FVisible := AReader.ReadBoolean;
  FFieldName := AReader.ReadString;
  FRecordOffset := AReader.ReadInteger;
End;

Class Function TLabelItem.CreateFromReader(AReader: TReader): TLabelItem;
Var
  LType: TLabelItemType;
Begin
  LType := TLabelItemType(AReader.ReadInteger);
  Case LType Of
    litText: Result := TLabelText.Create;
    litImage: Result := TLabelImage.Create;
    litBarcode: Result := TLabelBarcode.Create;
    litGrid: Result := TLabelGrid.Create;
    Else
      Result := nil;
  End;
  If Assigned(Result) Then
    Result.LoadFromReader(AReader);
End;

{ TLabelText }

Constructor TLabelText.Create;
Begin
  Inherited Create;
  FFont := TFont.Create;
  FText := 'Novo Texto';
  FAlignment := taLeftJustify;
  FVertAlignment := tlTop;
  FWordWrap := False;
  FWidth := 100;
  FHeight := 20;
End;

Destructor TLabelText.Destroy;
Begin
  FFont.Free;
  Inherited Destroy;
End;

Procedure TLabelText.SetText(AValue: String);
Begin
  If FText = AValue Then Exit;
  FText := AValue;
End;

Procedure TLabelText.Draw(ACanvas: TCanvas);
Var
  TS: TTextStyle;
Begin
  If NOT FVisible Then Exit;

  ACanvas.Font.Assign(FFont);

  TS := ACanvas.TextStyle;
  TS.Alignment := FAlignment;
  TS.Layout := FVertAlignment;
  TS.Wordbreak := FWordWrap;
  TS.SingleLine := NOT FWordWrap;
  TS.Clipping := True;

  ACanvas.TextRect(Rect(FLeft, FTop, FLeft + FWidth, FTop + FHeight), FLeft, FTop, FText, TS);

  If FSelected Then
  Begin
    ACanvas.Pen.Style := psDot;
    ACanvas.Pen.Color := clRed;
    ACanvas.Brush.Style := bsClear;
    ACanvas.Rectangle(FLeft - 2, FTop - 2, FLeft + FWidth + 2, FTop + FHeight + 2);
  End;
End;

Function TLabelText.Clone: TLabelItem;
Var
  NewItem: TLabelText;
Begin
  NewItem := TLabelText.Create;
  NewItem.Name := Name;
  NewItem.Left := Left;
  NewItem.Top := Top;
  NewItem.Width := Width;
  NewItem.Height := Height;
  NewItem.Text := Text;
  NewItem.Font.Assign(Font);
  NewItem.Alignment := Alignment;
  NewItem.VertAlignment := VertAlignment;
  NewItem.WordWrap := WordWrap;
  NewItem.Popup := Popup;
  NewItem.FieldName := FieldName;
  NewItem.RecordOffset := RecordOffset;
  Result := NewItem;
End;

Function TLabelText.GetItemType: TLabelItemType;
Begin
  Result := litText;
End;

Procedure TLabelText.SaveToWriter(AWriter: TWriter);
Begin
  AWriter.WriteInteger(Ord(GetItemType));
  Inherited SaveToWriter(AWriter);
  AWriter.WriteString(FText);
  AWriter.WriteInteger(Ord(FAlignment));
  // Font
  AWriter.WriteString(FFont.Name);
  AWriter.WriteInteger(FFont.Size);
  AWriter.WriteInteger(FFont.Color);
  AWriter.WriteInteger(Integer(FFont.Style));
  AWriter.WriteBoolean(FWordWrap);
  AWriter.WriteInteger(Ord(FVertAlignment));
End;

Procedure TLabelText.LoadFromReader(AReader: TReader);
Var
  LFontStyle: Integer;
Begin
  Inherited LoadFromReader(AReader);
  FText := AReader.ReadString;
  FAlignment := TAlignment(AReader.ReadInteger);
  // Font
  FFont.Name := AReader.ReadString;
  FFont.Size := AReader.ReadInteger;
  FFont.Color := TColor(AReader.ReadInteger);
  LFontStyle := AReader.ReadInteger;
  FFont.Style := TFontStyles(LFontStyle);
  Try
    FWordWrap := AReader.ReadBoolean;
  Except
    FWordWrap := False;
  End;

  If AReader.NextValue IN [vaInt8, vaInt16, vaInt32] Then
    FVertAlignment := TTextLayout(AReader.ReadInteger)
  Else
    FVertAlignment := tlTop;
End;

{ TLabelImage }

Constructor TLabelImage.Create;
Begin
  Inherited Create;
  FPicture := TPicture.Create;
  FStretch := True;
  FProportional := True;
End;

Destructor TLabelImage.Destroy;
Begin
  FPicture.Free;
  Inherited Destroy;
End;

Procedure TLabelImage.Draw(ACanvas: TCanvas);
Var
  R: TRect;
  ImgW, ImgH: Integer;
  Aspect: Double;
  TargetW, TargetH: Integer;
  OffsetX, OffsetY: Integer;
Begin
  If NOT FVisible Then Exit;

  If FPicture.Graphic <> nil Then
  Begin
    If FStretch Then
    Begin
      If FProportional Then
      Begin
        ImgW := FPicture.Graphic.Width;
        ImgH := FPicture.Graphic.Height;
        If (ImgW > 0) AND (ImgH > 0) Then
        Begin
          Aspect := ImgW / ImgH;
          If (FWidth / FHeight) > Aspect Then
          Begin
            TargetH := FHeight;
            TargetW := Round(TargetH * Aspect);
          End
          Else
          Begin
            TargetW := FWidth;
            TargetH := Round(TargetW / Aspect);
          End;
          OffsetX := (FWidth - TargetW) DIV 2;
          OffsetY := (FHeight - TargetH) DIV 2;
          R := Rect(FLeft + OffsetX, FTop + OffsetY, FLeft + OffsetX + TargetW, FTop + OffsetY + TargetH);
          ACanvas.StretchDraw(R, FPicture.Graphic);
        End;
      End
      Else
      Begin
        R := Rect(FLeft, FTop, FLeft + FWidth, FTop + FHeight);
        ACanvas.StretchDraw(R, FPicture.Graphic);
      End;
    End
    Else
      ACanvas.Draw(FLeft, FTop, FPicture.Graphic);
  End
  Else
  Begin
    ACanvas.Brush.Color := clSilver;
    ACanvas.Rectangle(FLeft, FTop, FLeft + FWidth, FTop + FHeight);
    ACanvas.TextOut(FLeft + 5, FTop + 5, 'Imagem');
  End;

  If FSelected Then
  Begin
    ACanvas.Pen.Style := psDot;
    ACanvas.Pen.Color := clRed;
    ACanvas.Brush.Style := bsClear;
    ACanvas.Rectangle(FLeft, FTop, FLeft + FWidth, FTop + FHeight);
  End;
End;

Function TLabelImage.Clone: TLabelItem;
Var
  NewItem: TLabelImage;
Begin
  NewItem := TLabelImage.Create;
  NewItem.Name := Name;
  NewItem.Left := Left;
  NewItem.Top := Top;
  NewItem.Width := Width;
  NewItem.Height := Height;
  NewItem.Picture.Assign(Picture);
  NewItem.Stretch := Stretch;
  NewItem.Proportional := Proportional;
  NewItem.Popup := Popup;
  NewItem.FieldName := FieldName;
  NewItem.RecordOffset := RecordOffset;
  Result := NewItem;
End;

Function TLabelImage.GetItemType: TLabelItemType;
Begin
  Result := litImage;
End;

Procedure TLabelImage.SaveToWriter(AWriter: TWriter);
Var
  MS: TMemoryStream;
Begin
  AWriter.WriteInteger(Ord(GetItemType));
  Inherited SaveToWriter(AWriter);
  AWriter.WriteBoolean(FStretch);
  AWriter.WriteBoolean(FProportional);
  If Assigned(FPicture.Graphic) AND (NOT FPicture.Graphic.Empty) Then
  Begin
    MS := TMemoryStream.Create;
    Try
      FPicture.Graphic.SaveToStream(MS);
      AWriter.WriteInteger(MS.Size);
      AWriter.Write(MS.Memory^, MS.Size);
    Finally
      MS.Free;
    End;
  End
  Else
    AWriter.WriteInteger(0);
End;

Procedure TLabelImage.LoadFromReader(AReader: TReader);
Var
  LSize: Integer;
  MS: TMemoryStream;
Begin
  Inherited LoadFromReader(AReader);
  // Compatibilidade com versões anteriores:
  // Se o próximo valor no stream não for booleano, significa que é um ficheiro antigo
  // que contém diretamente o tamanho da imagem (Integer).
  If AReader.NextValue IN [vaFalse, vaTrue] Then
  Begin
    FStretch := AReader.ReadBoolean;
    FProportional := AReader.ReadBoolean;
  End
  Else
  Begin
    FStretch := True;
    FProportional := True;
  End;

  LSize := AReader.ReadInteger;
  If LSize > 0 Then
  Begin
    MS := TMemoryStream.Create;
    Try
      MS.SetSize(LSize);
      AReader.Read(MS.Memory^, LSize);
      MS.Position := 0;
      FPicture.LoadFromStream(MS);
    Finally
      MS.Free;
    End;
  End;
End;

{ TLabelBarcode }

Constructor TLabelBarcode.Create;
Begin
  Inherited Create;
  FText := '123456789';
  FShowText := True;
  FKind := bkCode128;
  FModule := 2.0;
  FRatio := 2.0;
  UpdateDimensions;
End;

Procedure TLabelBarcode.SetRatio(AValue: Double);
Begin
  If FRatio <> AValue Then
  Begin
    FRatio := AValue;
    UpdateDimensions;
  End;
End;

Procedure TLabelBarcode.UpdateDimensions;
Var
  BC128: TBarcodeC128;
  BCEAN: TBarcodeEAN;
  BCQR: TBarcodeQR;
  TempText: String;
Begin
  TempText := FText;
  If TempText = '' Then TempText := '123456789';

  Try
    Case FKind Of
      bkCode128:
      Begin
        BC128 := TBarcodeC128.Create(nil);
        Try
          BC128.AutoSize := True;
          BC128.Text := TempText;
          BC128.Scale := Round(FModule);
          FWidth := BC128.Width;
        Finally
          BC128.Free;
        End;
      End;
      bkEAN13:
      Begin
        BCEAN := TBarcodeEAN.Create(nil);
        Try
          BCEAN.AutoSize := True;
          BCEAN.Text := TempText;
          BCEAN.Scale := Round(FModule);
          FWidth := BCEAN.Width;
        Finally
          BCEAN.Free;
        End;
      End;
      bkQR:
      Begin
        BCQR := TBarcodeQR.Create(nil);
        Try
          BCQR.AutoSize := True;
          BCQR.Text := TempText;
          BCQR.Scale := Round(FModule);
          FWidth := BCQR.Width;
          FHeight := BCQR.Height;
        Finally
          BCQR.Free;
        End;
      End;
      Else;
    End;

    If FWidth < 10 Then FWidth := 10;
    If FHeight < 10 Then FHeight := 10;
  Except
    SysUtils.Beep;
  End;
End;

Procedure TLabelBarcode.SetText(AValue: String);
Begin
  If FText <> AValue Then
  Begin
    FText := AValue;
    UpdateDimensions;
  End;
End;

Procedure TLabelBarcode.SetKind(AValue: TBarcodeKind);
Begin
  If FKind <> AValue Then
  Begin
    FKind := AValue;
    UpdateDimensions;
  End;
End;

Procedure TLabelBarcode.SetModule(AValue: Double);
Begin
  If FModule <> AValue Then
  Begin
    FModule := AValue;
    UpdateDimensions;
  End;
End;

Procedure TLabelBarcode.SetShowText(AValue: Boolean);
Begin
  FShowText := AValue;
End;

Procedure TLabelBarcode.SetWidth(AValue: Integer);
Var
  BaseWidth: Integer;
  NewModule: Double;
  BC128: TBarcodeC128;
  BCEAN: TBarcodeEAN;
  BCQR: TBarcodeQR;
  TempText: String;
  B: TBitmap;
  x, y: Integer;
  Found: Boolean;
Begin
  If AValue <= 0 Then Exit;

  TempText := FText;
  If TempText = '' Then TempText := '123456789';

  Try
    BaseWidth := 0;
    B := TBitmap.Create;
    Try
      B.PixelFormat := pf24bit;

      Case FKind Of
        bkCode128:
        Begin
          BC128 := TBarcodeC128.Create(nil);
          Try
            BC128.Text := TempText;
            BC128.ShowHumanReadableText := False;
            BC128.Scale := 1;
            BC128.AutoSize := True;

            B.SetSize(BC128.Width, BC128.Height);
            B.Canvas.Brush.Color := clWhite;
            B.Canvas.FillRect(0, 0, B.Width, B.Height);
            BC128.PaintOnCanvas(B.Canvas, Classes.Rect(0, 0, B.Width, B.Height));
          Finally
            BC128.Free;
          End;
        End;
        bkEAN13:
        Begin
          BCEAN := TBarcodeEAN.Create(nil);
          Try
            BCEAN.Text := TempText;
            BCEAN.ShowHumanReadableText := False;
            BCEAN.Scale := 1;
            BCEAN.AutoSize := True;

            B.SetSize(BCEAN.Width, BCEAN.Height);
            B.Canvas.Brush.Color := clWhite;
            B.Canvas.FillRect(0, 0, B.Width, B.Height);
            BCEAN.PaintOnCanvas(B.Canvas, Classes.Rect(0, 0, B.Width, B.Height));
          Finally
            BCEAN.Free;
          End;
        End;
        bkQR:
        Begin
          BCQR := TBarcodeQR.Create(nil);
          Try
            BCQR.Text := TempText;
            BCQR.Scale := 1;
            BCQR.AutoSize := True;

            B.SetSize(BCQR.Width, BCQR.Height);
            B.Canvas.Brush.Color := clWhite;
            B.Canvas.FillRect(0, 0, B.Width, B.Height);
            BCQR.PaintOnCanvas(B.Canvas, Classes.Rect(0, 0, B.Width, B.Height));
          Finally
            BCQR.Free;
          End;
        End;
        Else;
      End;

      BaseWidth := 1;
      Found := False;
      y := B.Height DIV 2;
      For x := B.Width - 1 Downto 0 Do
      Begin
        If B.Canvas.Pixels[x, y] <> clWhite Then
        Begin
          BaseWidth := x + 1;
          Found := True;
          Break;
        End;
      End;

      If NOT Found Then BaseWidth := B.Width;

    Finally
      B.Free;
    End;

    If BaseWidth > 0 Then
    Begin
      NewModule := AValue / BaseWidth;
      If NewModule < 0.1 Then NewModule := 0.1;

      FModule := NewModule;
      FWidth := AValue;
    End;
  Except
    SysUtils.Beep;
  End;
End;

Procedure TLabelBarcode.Draw(ACanvas: TCanvas);
Var
  R: TRect;
  BC128: TBarcodeC128;
  BCEAN: TBarcodeEAN;
  BCQR: TBarcodeQR;
  B: TBitmap;
  IntScale: Integer;
  DimTextH, DimBarH, DimTextW: Integer;
  RBar: TRect;
  RealWidth: Integer;
  TopY, BottomY: Integer;
  x, y: Integer;
  TargetFontSize: Integer;
  NaturalWidth: Integer;
  Found: Boolean;
  Line: Pbyte;
Begin
  If NOT FVisible Then Exit;

  R := Classes.Rect(FLeft, FTop, FLeft + FWidth, FTop + FHeight);

  ACanvas.Brush.Color := clWhite;
  ACanvas.Brush.Style := bsSolid;
  ACanvas.FillRect(R);

  Try
    IntScale := Trunc(FModule);
    If IntScale < 1 Then IntScale := 1;

    SetStretchBltMode(ACanvas.Handle, HALFTONE);
    SetBrushOrgEx(ACanvas.Handle, 0, 0, nil);

    B := TBitmap.Create;
    Try
      B.PixelFormat := pf24bit;

      TargetFontSize := Round(8 * FModule);
      If TargetFontSize > (FHeight DIV 4) Then TargetFontSize := FHeight DIV 4;
      If TargetFontSize < 6 Then TargetFontSize := 6;

      ACanvas.Font.Name := 'Arial';
      ACanvas.Font.Size := TargetFontSize;
      ACanvas.Brush.Style := bsClear;

      Case FKind Of
        bkCode128, bkEAN13:
        Begin
          DimTextH := 0;
          If FShowText Then
            DimTextH := ACanvas.TextHeight('Wg') + 2;

          DimBarH := FHeight - DimTextH;
          If DimBarH < 1 Then DimBarH := 1;

          RBar := Classes.Rect(FLeft, FTop, FLeft + FWidth, FTop + DimBarH);

          If FKind = bkCode128 Then
          Begin
            BC128 := TBarcodeC128.Create(nil);
            Try
              BC128.Text := FText;
              BC128.ShowHumanReadableText := False;
              BC128.Scale := IntScale;
              BC128.AutoSize := True;
              NaturalWidth := BC128.Width;

              B.SetSize(NaturalWidth, DimBarH);
              B.Canvas.Brush.Color := clWhite;
              B.Canvas.FillRect(0, 0, B.Width, B.Height);

              BC128.AutoSize := False;
              BC128.Width := NaturalWidth;
              BC128.Height := DimBarH;
              BC128.PaintOnCanvas(B.Canvas, Classes.Rect(0, 0, B.Width, B.Height));
            Finally
              BC128.Free;
            End;
          End
          Else
          Begin
            BCEAN := TBarcodeEAN.Create(nil);
            Try
              BCEAN.Text := FText;
              BCEAN.ShowHumanReadableText := False;
              BCEAN.Scale := IntScale;
              BCEAN.AutoSize := True;
              NaturalWidth := BCEAN.Width;

              B.SetSize(NaturalWidth, DimBarH);
              B.Canvas.Brush.Color := clWhite;
              B.Canvas.FillRect(0, 0, B.Width, B.Height);

              BCEAN.AutoSize := False;
              BCEAN.Width := NaturalWidth;
              BCEAN.Height := DimBarH;
              BCEAN.PaintOnCanvas(B.Canvas, Classes.Rect(0, 0, B.Width, B.Height));
            Finally
              BCEAN.Free;
            End;
          End;

          B.BeginUpdate;
          Try
            TopY := 0;
            BottomY := B.Height - 1;
            RealWidth := B.Width;
            Found := False;

            For y := 0 To B.Height - 1 Do
            Begin
              Line := B.ScanLine[y];
              For x := 0 To B.Width - 1 Do
              Begin
                If (Line[x * 3] < 255) OR (Line[x * 3 + 1] < 255) OR (Line[x * 3 + 2] < 255) Then
                Begin
                  TopY := y;
                  Found := True;
                  Break;
                End;
              End;
              If Found Then Break;
            End;

            If Found Then
            Begin
              For y := B.Height - 1 Downto TopY Do
              Begin
                Line := B.ScanLine[y];
                Found := False;
                For x := 0 To B.Width - 1 Do
                Begin
                  If (Line[x * 3] < 255) OR (Line[x * 3 + 1] < 255) OR (Line[x * 3 + 2] < 255) Then
                  Begin
                    Found := True;
                    Break;
                  End;
                End;
                If Found Then
                Begin
                  BottomY := y;
                  Break;
                End;
              End;

              For x := B.Width - 1 Downto 0 Do
              Begin
                Found := False;
                For y := TopY To BottomY Do
                Begin
                  Line := B.ScanLine[y];
                  If (Line[x * 3] < 255) OR (Line[x * 3 + 1] < 255) OR (Line[x * 3 + 2] < 255) Then
                  Begin
                    Found := True;
                    Break;
                  End;
                End;
                If Found Then
                Begin
                  RealWidth := x + 1;
                  Break;
                End;
              End;
            End;
          Finally
            B.EndUpdate;
          End;

          ACanvas.CopyRect(RBar, B.Canvas, Classes.Rect(0, TopY, RealWidth, BottomY + 1));

          If FShowText Then
          Begin
            DimTextW := ACanvas.TextWidth(FText);
            ACanvas.TextOut(FLeft + (FWidth - DimTextW) DIV 2, FTop + DimBarH, FText);
          End;
        End;

        bkQR:
        Begin
          BCQR := TBarcodeQR.Create(nil);
          Try
            BCQR.Text := FText;
            BCQR.Scale := IntScale;
            BCQR.AutoSize := True;

            B.SetSize(BCQR.Width, BCQR.Height);
            B.Canvas.Brush.Color := clWhite;
            B.Canvas.FillRect(0, 0, B.Width, B.Height);
            BCQR.PaintOnCanvas(B.Canvas, Classes.Rect(0, 0, B.Width, B.Height));
          Finally
            BCQR.Free;
          End;
          ACanvas.StretchDraw(R, B);
        End;
        Else;
      End;
    Finally
      B.Free;
    End;
  Except
    on E: Exception Do
    Begin
      ACanvas.Font.Color := clRed;
      ACanvas.Font.Size := 8;
      ACanvas.TextOut(FLeft, FTop, 'Erro');
    End;
  End;

  If FSelected Then
  Begin
    ACanvas.Pen.Style := psDot;
    ACanvas.Pen.Color := clRed;
    ACanvas.Brush.Style := bsClear;
    ACanvas.Rectangle(FLeft, FTop, FLeft + FWidth, FTop + FHeight);
  End;
End;

Function TLabelBarcode.Clone: TLabelItem;
Var
  NewItem: TLabelBarcode;
Begin
  NewItem := TLabelBarcode.Create;
  NewItem.Name := Name;
  NewItem.Left := Left;
  NewItem.Top := Top;
  NewItem.Width := Width;
  NewItem.Height := Height;
  NewItem.Text := Text;
  NewItem.ShowText := ShowText;
  NewItem.Kind := Kind;
  NewItem.Module := Module;
  NewItem.Ratio := Ratio;
  NewItem.Popup := Popup;
  NewItem.FieldName := FieldName;
  NewItem.RecordOffset := RecordOffset;
  Result := NewItem;
End;

Function TLabelBarcode.GetItemType: TLabelItemType;
Begin
  Result := litBarcode;
End;

Procedure TLabelBarcode.SaveToWriter(AWriter: TWriter);
Begin
  AWriter.WriteInteger(Ord(GetItemType));
  Inherited SaveToWriter(AWriter);
  AWriter.WriteString(FText);
  AWriter.WriteBoolean(FShowText);
  AWriter.WriteInteger(Ord(FKind));
  AWriter.WriteFloat(FModule);
  AWriter.WriteFloat(FRatio);
End;

Procedure TLabelBarcode.LoadFromReader(AReader: TReader);
Begin
  Inherited LoadFromReader(AReader);
  FText := AReader.ReadString;
  FShowText := AReader.ReadBoolean;
  FKind := TBarcodeKind(AReader.ReadInteger);
  FModule := AReader.ReadFloat;
  Try
    FRatio := AReader.ReadFloat;
  Except
    FRatio := 2.0;
  End;
  UpdateDimensions;
End;

{ TLabelGrid }

Constructor TLabelGrid.Create;
Begin
  Inherited Create;
  FRows := 2;
  FCols := 2;
  SetLength(FSelectedCells, 0);
  SetLength(FMergedCells, 0);
  UpdateCells;
  InitializeCellFormats;
  // Initialize uniform sizes
  SetLength(FColWidths, FCols);
  SetLength(FRowHeights, FRows);
  UpdateDimensions;
End;

Procedure TLabelGrid.SetWidth(AValue: Integer);
Var
  i: Integer;
  Ratio: Double;
Begin
  If (FWidth > 0) AND (AValue <> FWidth) Then
  Begin
    Ratio := AValue / FWidth;
    For i := 0 To FCols - 1 Do
      FColWidths[i] := Round(FColWidths[i] * Ratio);
  End;
  Inherited SetWidth(AValue);
End;

Procedure TLabelGrid.SetHeight(AValue: Integer);
Var
  i: Integer;
  Ratio: Double;
Begin
  If (FHeight > 0) AND (AValue <> FHeight) Then
  Begin
    Ratio := AValue / FHeight;
    For i := 0 To FRows - 1 Do
      FRowHeights[i] := Round(FRowHeights[i] * Ratio);
  End;
  Inherited SetHeight(AValue);
End;

Destructor TLabelGrid.Destroy;
Var
  r, c: Integer;
Begin
  For r := 0 To Length(FCells) - 1 Do
    For c := 0 To Length(FCells[r]) - 1 Do
      If Assigned(FCells[r, c]) Then FCells[r, c].Free;
  Inherited Destroy;
End;

Procedure TLabelGrid.UpdateCells;
Var
  OldCells: Array Of Array Of TLabelItem;
  r, c: Integer;
  OldRows, OldCols: Integer;
Begin
  OldCells := FCells;
  OldRows := Length(OldCells);
  If OldRows > 0 Then OldCols := Length(OldCells[0])
  Else
    OldCols := 0;

  SetLength(FCells, FRows, FCols);

  For r := 0 To OldRows - 1 Do
  Begin
    For c := 0 To OldCols - 1 Do
    Begin
      If (r < FRows) AND (c < FCols) Then
        FCells[r, c] := OldCells[r, c]
      Else
        OldCells[r, c].Free;
    End;
  End;

  For r := 0 To FRows - 1 Do
    For c := 0 To FCols - 1 Do
      If (r >= OldRows) OR (c >= OldCols) Then
        FCells[r, c] := nil;

  // Update dimension arrays
  SetLength(FColWidths, FCols);
  SetLength(FRowHeights, FRows);
  UpdateDimensions;
End;

Procedure TLabelGrid.UpdateDimensions;
Var
  i: Integer;
Begin
  If FCols > 0 Then
    For i := 0 To FCols - 1 Do FColWidths[i] := FWidth DIV FCols;
  If FRows > 0 Then
    For i := 0 To FRows - 1 Do FRowHeights[i] := FHeight DIV FRows;
End;

Procedure TLabelGrid.UpdateFormats;
Var
  OldFmt: Array Of Array Of TCellFormat;
  OldRows, OldCols: Integer;
  r, c: Integer;
Begin
  OldFmt := FCellFormats;
  OldRows := Length(OldFmt);
  If OldRows > 0 Then OldCols := Length(OldFmt[0])
  Else
    OldCols := 0;

  SetLength(FCellFormats, FRows, FCols);

  For r := 0 To FRows - 1 Do
  Begin
    For c := 0 To FCols - 1 Do
    Begin
      If (r < OldRows) AND (c < OldCols) Then
        FCellFormats[r, c] := OldFmt[r, c]
      Else
      Begin
        FCellFormats[r, c].BackColor := clWhite;
        FCellFormats[r, c].BorderColor := clBlack;
        FCellFormats[r, c].BorderVisible := [beLeft, beTop, beRight, beBottom];
      End;
    End;
  End;
End;

Procedure TLabelGrid.InitializeCellFormats;
Var
  r, c: Integer;
Begin
  SetLength(FCellFormats, FRows, FCols);
  For r := 0 To FRows - 1 Do
  Begin
    For c := 0 To FCols - 1 Do
    Begin
      FCellFormats[r, c].BackColor := clWhite;
      FCellFormats[r, c].BorderColor := clBlack;
      FCellFormats[r, c].BorderVisible := [beLeft, beTop, beRight, beBottom];
    End;
  End;
End;

Procedure TLabelGrid.SetRows(AValue: Integer);
Begin
  If AValue < 1 Then AValue := 1;
  If FRows <> AValue Then
  Begin
    FRows := AValue;
    UpdateCells;
    UpdateFormats;
    NormalizeMerges;
    ClearSelection;
  End;
End;

Procedure TLabelGrid.SetCols(AValue: Integer);
Begin
  If AValue < 1 Then AValue := 1;
  If FCols <> AValue Then
  Begin
    FCols := AValue;
    UpdateCells;
    UpdateFormats;
    NormalizeMerges;
    ClearSelection;
  End;
End;

Procedure TLabelGrid.NormalizeMerges;
Var
  i, j: Integer;
Begin
  For i := Length(FMergedCells) - 1 Downto 0 Do
  Begin
    With FMergedCells[i] Do
    Begin
      If (Row < 0) OR (Col < 0) OR (RowSpan < 1) OR (ColSpan < 1) OR (Row + RowSpan > FRows) OR (Col + ColSpan > FCols) Then
      Begin
        // Remove invalid merge
        For j := i To Length(FMergedCells) - 2 Do
          FMergedCells[j] := FMergedCells[j + 1];
        SetLength(FMergedCells, Length(FMergedCells) - 1);
      End;
    End;
  End;
End;

// Selection Methods

Procedure TLabelGrid.ClearSelection;
Begin
  SetLength(FSelectedCells, 0);
End;

Procedure TLabelGrid.SelectCell(ARow, ACol: Integer; AAdd: Boolean = False);
Var
  i: Integer;
Begin
  If (ARow < 0) OR (ARow >= FRows) OR (ACol < 0) OR (ACol >= FCols) Then Exit;

  If NOT AAdd Then
    ClearSelection;

  // Check if already selected
  For i := 0 To Length(FSelectedCells) - 1 Do
    If (FSelectedCells[i].Y = ARow) AND (FSelectedCells[i].X = ACol) Then Exit;

  SetLength(FSelectedCells, Length(FSelectedCells) + 1);
  FSelectedCells[High(FSelectedCells)] := Point(ACol, ARow);
End;

Procedure TLabelGrid.ToggleCellSelection(ARow, ACol: Integer);
Var
  i, j: Integer;
  Found: Boolean;
Begin
  If (ARow < 0) OR (ARow >= FRows) OR (ACol < 0) OR (ACol >= FCols) Then Exit;

  Found := False;
  For i := 0 To Length(FSelectedCells) - 1 Do
  Begin
    If (FSelectedCells[i].Y = ARow) AND (FSelectedCells[i].X = ACol) Then
    Begin
      // Remove
      For j := i To Length(FSelectedCells) - 2 Do
        FSelectedCells[j] := FSelectedCells[j + 1];
      SetLength(FSelectedCells, Length(FSelectedCells) - 1);
      Found := True;
      Break;
    End;
  End;

  If NOT Found Then
  Begin
    SetLength(FSelectedCells, Length(FSelectedCells) + 1);
    FSelectedCells[High(FSelectedCells)] := Point(ACol, ARow);
  End;
End;

Procedure TLabelGrid.SelectRange(ARow1, ACol1, ARow2, ACol2: Integer; AAdd: Boolean = False);
Var
  r, c: Integer;
  MinR, MaxR, MinC, MaxC: Integer;
Begin
  If NOT AAdd Then ClearSelection;

  MinR := ARow1;
  If ARow2 < MinR Then MinR := ARow2;
  MaxR := ARow1;
  If ARow2 > MaxR Then MaxR := ARow2;
  MinC := ACol1;
  If ACol2 < MinC Then MinC := ACol2;
  MaxC := ACol1;
  If ACol2 > MaxC Then MaxC := ACol2;

  For r := MinR To MaxR Do
    For c := MinC To MaxC Do
      SelectCell(r, c, True);
End;

Function TLabelGrid.IsSelected(ARow, ACol: Integer): Boolean;
Var
  i: Integer;
Begin
  Result := False;
  For i := 0 To Length(FSelectedCells) - 1 Do
    If (FSelectedCells[i].Y = ARow) AND (FSelectedCells[i].X = ACol) Then
    Begin
      Result := True;
      Exit;
    End;
End;

Function TLabelGrid.GetSelectionCount: Integer;
Begin
  Result := Length(FSelectedCells);
End;

Function TLabelGrid.GetFirstSelectedCell: TPoint;
Begin
  If Length(FSelectedCells) > 0 Then
    Result := FSelectedCells[0]
  Else
    Result := Point(-1, -1);
End;

Function TLabelGrid.CanGroupSelected: Boolean;
Var
  MinR, MaxR, MinC, MaxC: Integer;
  i: Integer;
  Count: Integer;
Begin
  Result := False;
  Count := Length(FSelectedCells);
  If Count < 2 Then Exit;

  // Calculate bounds
  MinR := FRows;
  MaxR := -1;
  MinC := FCols;
  MaxC := -1;

  For i := 0 To Count - 1 Do
  Begin
    If FSelectedCells[i].Y < MinR Then MinR := FSelectedCells[i].Y;
    If FSelectedCells[i].Y > MaxR Then MaxR := FSelectedCells[i].Y;
    If FSelectedCells[i].X < MinC Then MinC := FSelectedCells[i].X;
    If FSelectedCells[i].X > MaxC Then MaxC := FSelectedCells[i].X;
  End;

  // Check if selection forms a rectangle and matches count
  If ((MaxR - MinR + 1) * (MaxC - MinC + 1)) <> Count Then Exit;

  Result := True;
End;

Function TLabelGrid.CanUngroupSelected: Boolean;
Var
  i: Integer;
  MergeInfo: TMergedCell;
Begin
  Result := False;
  For i := 0 To Length(FSelectedCells) - 1 Do
  Begin
    MergeInfo.Row := -1;
    If IsCellMerged(FSelectedCells[i].Y, FSelectedCells[i].X, MergeInfo) Then
    Begin
      Result := True;
      Exit;
    End;
  End;
End;

Procedure TLabelGrid.GroupSelected;
Var
  MinR, MaxR, MinC, MaxC: Integer;
  i: Integer;
Begin
  If NOT CanGroupSelected Then Exit;

  // Calculate bounds
  MinR := FRows;
  MaxR := -1;
  MinC := FCols;
  MaxC := -1;

  For i := 0 To Length(FSelectedCells) - 1 Do
  Begin
    If FSelectedCells[i].Y < MinR Then MinR := FSelectedCells[i].Y;
    If FSelectedCells[i].Y > MaxR Then MaxR := FSelectedCells[i].Y;
    If FSelectedCells[i].X < MinC Then MinC := FSelectedCells[i].X;
    If FSelectedCells[i].X > MaxC Then MaxC := FSelectedCells[i].X;
  End;

  MergeCells(MinR, MinC, MaxR - MinR + 1, MaxC - MinC + 1);

  // Select the merged cell (top-left)
  ClearSelection;
  SelectCell(MinR, MinC);
End;

Procedure TLabelGrid.UngroupSelected;
Var
  i: Integer;
Begin
  For i := 0 To Length(FSelectedCells) - 1 Do
    UnmergeCells(FSelectedCells[i].Y, FSelectedCells[i].X);
End;

Procedure TLabelGrid.SetCellItem(ARow, ACol: Integer; AItem: TLabelItem);
Begin
  If (ARow >= 0) AND (ARow < FRows) AND (ACol >= 0) AND (ACol < FCols) Then
  Begin
    If Assigned(FCells[ARow, ACol]) Then FCells[ARow, ACol].Free;
    FCells[ARow, ACol] := AItem;
  End;
End;

Function TLabelGrid.GetCellItem(ARow, ACol: Integer): TLabelItem;
Begin
  If (ARow >= 0) AND (ARow < FRows) AND (ACol >= 0) AND (ACol < FCols) Then
    Result := FCells[ARow, ACol]
  Else
    Result := nil;
End;

Function TLabelGrid.FindChildAt(AX, AY: Integer): TLabelItem;
Var
  LRow, LCol: Integer;
Begin
  Result := nil;
  LRow := -1;
  LCol := -1;
  If GetCellAt(AX, AY, LRow, LCol) Then
    Result := GetCellItem(LRow, LCol);
End;

// Cell formatting methods

Procedure TLabelGrid.SetCellBackColor(ARow, ACol: Integer; AColor: TColor);
Begin
  If (ARow >= 0) AND (ARow < FRows) AND (ACol >= 0) AND (ACol < FCols) Then
    FCellFormats[ARow, ACol].BackColor := AColor;
End;

Function TLabelGrid.GetCellBackColor(ARow, ACol: Integer): TColor;
Begin
  If (ARow >= 0) AND (ARow < FRows) AND (ACol >= 0) AND (ACol < FCols) Then
    Result := FCellFormats[ARow, ACol].BackColor
  Else
    Result := clWhite;
End;

Procedure TLabelGrid.SetCellBorderColor(ARow, ACol: Integer; AColor: TColor);
Begin
  If (ARow >= 0) AND (ARow < FRows) AND (ACol >= 0) AND (ACol < FCols) Then
    FCellFormats[ARow, ACol].BorderColor := AColor;
End;

Function TLabelGrid.GetCellBorderColor(ARow, ACol: Integer): TColor;
Begin
  If (ARow >= 0) AND (ARow < FRows) AND (ACol >= 0) AND (ACol < FCols) Then
    Result := FCellFormats[ARow, ACol].BorderColor
  Else
    Result := clBlack;
End;

Procedure TLabelGrid.SetCellBorderVisible(ARow, ACol: Integer; AEdges: TBorderEdges);
Begin
  If (ARow >= 0) AND (ARow < FRows) AND (ACol >= 0) AND (ACol < FCols) Then
    FCellFormats[ARow, ACol].BorderVisible := AEdges;
End;

Procedure TLabelGrid.SetEdgeVisibility(ARow, ACol: Integer; AEdge: TBorderEdge; AVisible: Boolean);
Begin
  If (ARow < 0) OR (ARow >= FRows) OR (ACol < 0) OR (ACol >= FCols) Then Exit;

  // Update current cell
  If AVisible Then
    Include(FCellFormats[ARow, ACol].BorderVisible, AEdge)
  Else
    Exclude(FCellFormats[ARow, ACol].BorderVisible, AEdge);

  // Update adjacent cell
  Case AEdge Of
    beLeft:
    Begin
      If ACol > 0 Then
      Begin
        If AVisible Then
          Include(FCellFormats[ARow, ACol - 1].BorderVisible, beRight)
        Else
          Exclude(FCellFormats[ARow, ACol - 1].BorderVisible, beRight);
      End;
    End;
    beTop:
    Begin
      If ARow > 0 Then
      Begin
        If AVisible Then
          Include(FCellFormats[ARow - 1, ACol].BorderVisible, beBottom)
        Else
          Exclude(FCellFormats[ARow - 1, ACol].BorderVisible, beBottom);
      End;
    End;
    beRight:
    Begin
      If ACol < FCols - 1 Then
      Begin
        If AVisible Then
          Include(FCellFormats[ARow, ACol + 1].BorderVisible, beLeft)
        Else
          Exclude(FCellFormats[ARow, ACol + 1].BorderVisible, beLeft);
      End;
    End;
    beBottom:
    Begin
      If ARow < FRows - 1 Then
      Begin
        If AVisible Then
          Include(FCellFormats[ARow + 1, ACol].BorderVisible, beTop)
        Else
          Exclude(FCellFormats[ARow + 1, ACol].BorderVisible, beTop);
      End;
    End;
  End;
End;

Function TLabelGrid.GetCellBorderVisible(ARow, ACol: Integer): TBorderEdges;
Begin
  If (ARow >= 0) AND (ARow < FRows) AND (ACol >= 0) AND (ACol < FCols) Then
    Result := FCellFormats[ARow, ACol].BorderVisible
  Else
    Result := [beLeft, beTop, beRight, beBottom];
End;

// Merge operations

Function TLabelGrid.IsCellMerged(ARow, ACol: Integer; Var AMergeInfo: TMergedCell): Boolean;
Var
  i: Integer;
Begin
  Result := False;
  AMergeInfo.Row := -1;
  AMergeInfo.Col := -1;
  AMergeInfo.RowSpan := 0;
  AMergeInfo.ColSpan := 0;
  For i := 0 To Length(FMergedCells) - 1 Do
  Begin
    If (ARow >= FMergedCells[i].Row) AND (ARow < FMergedCells[i].Row + FMergedCells[i].RowSpan) AND (ACol >= FMergedCells[i].Col) AND
      (ACol < FMergedCells[i].Col + FMergedCells[i].ColSpan) Then
    Begin
      AMergeInfo := FMergedCells[i];
      Result := True;
      Exit;
    End;
  End;
End;

Function TLabelGrid.GetActualCell(ARow, ACol: Integer; Var AActualRow, AActualCol: Integer): Boolean;
Var
  MergeInfo: TMergedCell;
Begin
  MergeInfo.Row := -1;
  If IsCellMerged(ARow, ACol, MergeInfo) Then
  Begin
    AActualRow := MergeInfo.Row;
    AActualCol := MergeInfo.Col;
    Result := True;
  End
  Else
  Begin
    AActualRow := ARow;
    AActualCol := ACol;
    Result := False;
  End;
End;

Procedure TLabelGrid.MergeCells(AStartRow, AStartCol, ARowSpan, AColSpan: Integer);
Var
  NewMerge: TMergedCell;
  Len: Integer;
  r, c: Integer;
Begin
  // Validate parameters
  If (AStartRow < 0) OR (AStartRow >= FRows) OR (AStartCol < 0) OR (AStartCol >= FCols) Then Exit;
  If (ARowSpan < 1) OR (AColSpan < 1) Then Exit;
  If (AStartRow + ARowSpan > FRows) OR (AStartCol + AColSpan > FCols) Then Exit;

  // Remove any existing merges that overlap ANY cell in the new range
  For r := AStartRow To AStartRow + ARowSpan - 1 Do
    For c := AStartCol To AStartCol + AColSpan - 1 Do
      UnmergeCells(r, c);

  // Add new merge
  NewMerge.Row := AStartRow;
  NewMerge.Col := AStartCol;
  NewMerge.RowSpan := ARowSpan;
  NewMerge.ColSpan := AColSpan;

  Len := Length(FMergedCells);
  SetLength(FMergedCells, Len + 1);
  FMergedCells[Len] := NewMerge;
End;

Procedure TLabelGrid.UnmergeCells(ARow, ACol: Integer);
Var
  i, j: Integer;
Begin
  For i := Length(FMergedCells) - 1 Downto 0 Do
  Begin
    // Use explicit references to avoid name collision with parameters Row/Col
    If (ARow >= FMergedCells[i].Row) AND (ARow < FMergedCells[i].Row + FMergedCells[i].RowSpan) AND (ACol >= FMergedCells[i].Col) AND
      (ACol < FMergedCells[i].Col + FMergedCells[i].ColSpan) Then
    Begin
      // Remove this merge
      For j := i To Length(FMergedCells) - 2 Do
        FMergedCells[j] := FMergedCells[j + 1];
      SetLength(FMergedCells, Length(FMergedCells) - 1);
      Exit;
    End;
  End;
End;

Procedure TLabelGrid.ClearAllMerges;
Begin
  SetLength(FMergedCells, 0);
End;

Procedure TLabelGrid.InsertRow(AIndex: Integer);
Var
  r, c, i: Integer;
Begin
  If (AIndex < 0) OR (AIndex > FRows) Then Exit;

  SetLength(FCells, FRows + 1, FCols);
  SetLength(FCellFormats, FRows + 1, FCols);
  SetLength(FRowHeights, FRows + 1);

  // Shift down
  For r := FRows Downto AIndex + 1 Do
  Begin
    FRowHeights[r] := FRowHeights[r - 1];
    For c := 0 To FCols - 1 Do
    Begin
      FCells[r, c] := FCells[r - 1, c];
      FCellFormats[r, c] := FCellFormats[r - 1, c];
    End;
  End;

  // Clear new row
  FRowHeights[AIndex] := FHeight DIV (FRows + 1);
  For c := 0 To FCols - 1 Do
  Begin
    FCells[AIndex, c] := nil;
    FCellFormats[AIndex, c].BackColor := clWhite;
    FCellFormats[AIndex, c].BorderColor := clBlack;
    FCellFormats[AIndex, c].BorderVisible := [beLeft, beTop, beRight, beBottom];
  End;

  Inc(FRows);

  // Update Merges
  For i := 0 To Length(FMergedCells) - 1 Do
  Begin
    If FMergedCells[i].Row >= AIndex Then
      Inc(FMergedCells[i].Row)
    Else If (FMergedCells[i].Row < AIndex) AND (FMergedCells[i].Row + FMergedCells[i].RowSpan > AIndex) Then
      Inc(FMergedCells[i].RowSpan);
  End;

  // Update Selection
  For i := 0 To Length(FSelectedCells) - 1 Do
  Begin
    If FSelectedCells[i].Y >= AIndex Then
      Inc(FSelectedCells[i].Y);
  End;
End;

Procedure TLabelGrid.InsertCol(AIndex: Integer);
Var
  r, c, i: Integer;
Begin
  If (AIndex < 0) OR (AIndex > FCols) Then Exit;

  SetLength(FCells, FRows, FCols + 1);
  SetLength(FCellFormats, FRows, FCols + 1);
  SetLength(FColWidths, FCols + 1);

  // Shift right
  For r := 0 To FRows - 1 Do
  Begin
    For c := FCols Downto AIndex + 1 Do
    Begin
      If r = 0 Then FColWidths[c] := FColWidths[c - 1];
      FCells[r, c] := FCells[r, c - 1];
      FCellFormats[r, c] := FCellFormats[r, c - 1];
    End;

    // Clear new col
    If r = 0 Then FColWidths[AIndex] := FWidth DIV (FCols + 1);
    FCells[r, AIndex] := nil;
    FCellFormats[r, AIndex].BackColor := clWhite;
    FCellFormats[r, AIndex].BorderColor := clBlack;
    FCellFormats[r, AIndex].BorderVisible := [beLeft, beTop, beRight, beBottom];
  End;

  Inc(FCols);

  // Update Merges
  For i := 0 To Length(FMergedCells) - 1 Do
  Begin
    If FMergedCells[i].Col >= AIndex Then
      Inc(FMergedCells[i].Col)
    Else If (FMergedCells[i].Col < AIndex) AND (FMergedCells[i].Col + FMergedCells[i].ColSpan > AIndex) Then
      Inc(FMergedCells[i].ColSpan);
  End;

  // Update Selection
  For i := 0 To Length(FSelectedCells) - 1 Do
  Begin
    If FSelectedCells[i].X >= AIndex Then
      Inc(FSelectedCells[i].X);
  End;
End;

Procedure TLabelGrid.DeleteRow(AIndex: Integer);
Var
  r, c, i, j: Integer;
Begin
  If (AIndex < 0) OR (AIndex >= FRows) Then Exit;
  If FRows <= 1 Then Exit; // Don't delete last row

  // Free items in deleted row
  For c := 0 To FCols - 1 Do
    If Assigned(FCells[AIndex, c]) Then FCells[AIndex, c].Free;

  // Shift up
  For r := AIndex To FRows - 2 Do
  Begin
    FRowHeights[r] := FRowHeights[r + 1];
    For c := 0 To FCols - 1 Do
    Begin
      FCells[r, c] := FCells[r + 1, c];
      FCellFormats[r, c] := FCellFormats[r + 1, c];
    End;
  End;

  SetLength(FCells, FRows - 1, FCols);
  SetLength(FCellFormats, FRows - 1, FCols);
  SetLength(FRowHeights, FRows - 1);
  Dec(FRows);

  // Update Merges
  For i := Length(FMergedCells) - 1 Downto 0 Do
  Begin
    If (FMergedCells[i].Row = AIndex) AND (FMergedCells[i].RowSpan = 1) Then
    Begin
      // Remove
      For j := i To Length(FMergedCells) - 2 Do FMergedCells[j] := FMergedCells[j + 1];
      SetLength(FMergedCells, Length(FMergedCells) - 1);
    End
    Else If (FMergedCells[i].Row <= AIndex) AND (FMergedCells[i].Row + FMergedCells[i].RowSpan > AIndex) Then
    Begin
      Dec(FMergedCells[i].RowSpan);
    End
    Else If FMergedCells[i].Row > AIndex Then
    Begin
      Dec(FMergedCells[i].Row);
    End;
  End;

  // Update Selection
  For i := Length(FSelectedCells) - 1 Downto 0 Do
  Begin
    If FSelectedCells[i].Y = AIndex Then
    Begin
      // Remove selection
      For j := i To Length(FSelectedCells) - 2 Do FSelectedCells[j] := FSelectedCells[j + 1];
      SetLength(FSelectedCells, Length(FSelectedCells) - 1);
    End
    Else If FSelectedCells[i].Y > AIndex Then
    Begin
      Dec(FSelectedCells[i].Y);
    End;
  End;
End;

Procedure TLabelGrid.DeleteCol(AIndex: Integer);
Var
  r, c, i, j: Integer;
Begin
  If (AIndex < 0) OR (AIndex >= FCols) Then Exit;
  If FCols <= 1 Then Exit; // Don't delete last col

  // Free items in deleted col
  For r := 0 To FRows - 1 Do
    If Assigned(FCells[r, AIndex]) Then FCells[r, AIndex].Free;

  // Shift left
  For r := 0 To FRows - 1 Do
  Begin
    For c := AIndex To FCols - 2 Do
    Begin
      If r = 0 Then FColWidths[c] := FColWidths[c + 1];
      FCells[r, c] := FCells[r, c + 1];
      FCellFormats[r, c] := FCellFormats[r, c + 1];
    End;
  End;

  SetLength(FCells, FRows, FCols - 1);
  SetLength(FCellFormats, FRows, FCols - 1);
  SetLength(FColWidths, FCols - 1);
  Dec(FCols);

  // Update Merges
  For i := Length(FMergedCells) - 1 Downto 0 Do
  Begin
    If (FMergedCells[i].Col = AIndex) AND (FMergedCells[i].ColSpan = 1) Then
    Begin
      // Remove
      For j := i To Length(FMergedCells) - 2 Do FMergedCells[j] := FMergedCells[j + 1];
      SetLength(FMergedCells, Length(FMergedCells) - 1);
    End
    Else If (FMergedCells[i].Col <= AIndex) AND (FMergedCells[i].Col + FMergedCells[i].ColSpan > AIndex) Then
    Begin
      Dec(FMergedCells[i].ColSpan);
    End
    Else If FMergedCells[i].Col > AIndex Then
    Begin
      Dec(FMergedCells[i].Col);
    End;
  End;

  // Update Selection
  For i := Length(FSelectedCells) - 1 Downto 0 Do
  Begin
    If FSelectedCells[i].X = AIndex Then
    Begin
      // Remove selection
      For j := i To Length(FSelectedCells) - 2 Do FSelectedCells[j] := FSelectedCells[j + 1];
      SetLength(FSelectedCells, Length(FSelectedCells) - 1);
    End
    Else If FSelectedCells[i].X > AIndex Then
    Begin
      Dec(FSelectedCells[i].X);
    End;
  End;
End;

Procedure TLabelGrid.Draw(ACanvas: TCanvas);
Var
  r, c: Integer;
  X, Y, CX, CY, CW, CH: Integer;
  Item: TLabelItem;
  MergeInfo: TMergedCell;
  IsMerged: Boolean;
  CellRect: TRect;
  Format: TCellFormat;
  DrawnCells: Array Of Array Of Boolean;
  i, j: Integer;
  IsPrinting: Boolean;
Begin
  DrawnCells := nil;

  If NOT FVisible Then Exit;

  IsPrinting := (GetDeviceCaps(ACanvas.Handle, TECHNOLOGY) = DT_RASPRINTER);

  // Initialize drawn cells tracker
  SetLength(DrawnCells, FRows, FCols);
  For r := 0 To FRows - 1 Do
    For c := 0 To FCols - 1 Do
      DrawnCells[r, c] := False;

  // Draw cells with background colors and content
  For r := 0 To FRows - 1 Do
  Begin
    For c := 0 To FCols - 1 Do
    Begin
      If DrawnCells[r, c] Then Continue;

      // Initialize defaults to avoid uninitialized warnings
      Item := nil;
      Format := FCellFormats[r, c];

      MergeInfo.Row := -1;
      IsMerged := IsCellMerged(r, c, MergeInfo);

      If IsMerged Then
      Begin
        // Draw merged cell
        CX := FLeft;
        For i := 0 To MergeInfo.Col - 1 Do CX := CX + FColWidths[i];
        CY := FTop;
        For i := 0 To MergeInfo.Row - 1 Do CY := CY + FRowHeights[i];

        CW := 0;
        For i := MergeInfo.Col To MergeInfo.Col + MergeInfo.ColSpan - 1 Do CW := CW + FColWidths[i];
        CH := 0;
        For i := MergeInfo.Row To MergeInfo.Row + MergeInfo.RowSpan - 1 Do CH := CH + FRowHeights[i];

        // Mark all cells in merge as drawn
        For i := MergeInfo.Row To MergeInfo.Row + MergeInfo.RowSpan - 1 Do
          For j := MergeInfo.Col To MergeInfo.Col + MergeInfo.ColSpan - 1 Do
            If (i < FRows) AND (j < FCols) Then
              DrawnCells[i, j] := True;

        // Use format from top-left cell of merge
        Format := FCellFormats[MergeInfo.Row, MergeInfo.Col];
        // Use content from top-left cell of merge
        Item := FCells[MergeInfo.Row, MergeInfo.Col];
      End
      Else
      Begin
        // Draw single cell
        CX := FLeft;
        For i := 0 To c - 1 Do CX := CX + FColWidths[i];
        CY := FTop;
        For i := 0 To r - 1 Do CY := CY + FRowHeights[i];
        CW := FColWidths[c];
        CH := FRowHeights[r];
        DrawnCells[r, c] := True;
        // Format already initialized from current cell
        // Use content from this cell
        Item := FCells[r, c];
      End;

      CellRect := Rect(CX, CY, CX + CW, CY + CH);

      // Draw background
      ACanvas.Brush.Color := Format.BackColor;
      ACanvas.Brush.Style := bsSolid;
      ACanvas.FillRect(CellRect);

      // Draw cell content if exists
      If Assigned(Item) Then
      Begin
        Item.Left := CX;
        Item.Top := CY;
        Item.Width := CW;
        Item.Height := CH;
        Item.Draw(ACanvas);
      End;
    End;
  End;

  // Draw borders
  ACanvas.Brush.Style := bsClear;
  For r := 0 To FRows - 1 Do
  Begin
    For c := 0 To FCols - 1 Do
    Begin
      // Initialize defaults
      Format := FCellFormats[r, c];

      MergeInfo.Row := -1;
      IsMerged := IsCellMerged(r, c, MergeInfo);

      If IsMerged Then
      Begin
        // Only draw borders for top-left cell of merge
        If (r <> MergeInfo.Row) OR (c <> MergeInfo.Col) Then Continue;

        CX := FLeft;
        For i := 0 To MergeInfo.Col - 1 Do CX := CX + FColWidths[i];
        CY := FTop;
        For i := 0 To MergeInfo.Row - 1 Do CY := CY + FRowHeights[i];

        CW := 0;
        For i := MergeInfo.Col To MergeInfo.Col + MergeInfo.ColSpan - 1 Do CW := CW + FColWidths[i];
        CH := 0;
        For i := MergeInfo.Row To MergeInfo.Row + MergeInfo.RowSpan - 1 Do CH := CH + FRowHeights[i];

        Format := FCellFormats[MergeInfo.Row, MergeInfo.Col];
      End
      Else
      Begin
        CX := FLeft;
        For i := 0 To c - 1 Do CX := CX + FColWidths[i];
        CY := FTop;
        For i := 0 To r - 1 Do CY := CY + FRowHeights[i];
        CW := FColWidths[c];
        CH := FRowHeights[r];
        // Format already initialized from current cell
      End;

      // Pen setup is now handled per edge to support dashed lines for hidden borders

      // Draw individual borders

      // Left
      If beLeft IN Format.BorderVisible Then
      Begin
        ACanvas.Pen.Color := Format.BorderColor;
        ACanvas.Pen.Style := psSolid;
        ACanvas.MoveTo(CX, CY);
        ACanvas.LineTo(CX, CY + CH);
      End
      Else If NOT IsPrinting Then
      Begin
        ACanvas.Pen.Color := clSilver;
        ACanvas.Pen.Style := psDot;
        ACanvas.MoveTo(CX, CY);
        ACanvas.LineTo(CX, CY + CH);
      End;

      // Top
      If beTop IN Format.BorderVisible Then
      Begin
        ACanvas.Pen.Color := Format.BorderColor;
        ACanvas.Pen.Style := psSolid;
        ACanvas.MoveTo(CX, CY);
        ACanvas.LineTo(CX + CW, CY);
      End
      Else If NOT IsPrinting Then
      Begin
        ACanvas.Pen.Color := clSilver;
        ACanvas.Pen.Style := psDot;
        ACanvas.MoveTo(CX, CY);
        ACanvas.LineTo(CX + CW, CY);
      End;

      // Right
      If beRight IN Format.BorderVisible Then
      Begin
        ACanvas.Pen.Color := Format.BorderColor;
        ACanvas.Pen.Style := psSolid;
        ACanvas.MoveTo(CX + CW, CY);
        ACanvas.LineTo(CX + CW, CY + CH);
      End
      Else If NOT IsPrinting Then
      Begin
        ACanvas.Pen.Color := clSilver;
        ACanvas.Pen.Style := psDot;
        ACanvas.MoveTo(CX + CW, CY);
        ACanvas.LineTo(CX + CW, CY + CH);
      End;

      // Bottom
      If beBottom IN Format.BorderVisible Then
      Begin
        ACanvas.Pen.Color := Format.BorderColor;
        ACanvas.Pen.Style := psSolid;
        ACanvas.MoveTo(CX, CY + CH);
        ACanvas.LineTo(CX + CW, CY + CH);
      End
      Else If NOT IsPrinting Then
      Begin
        ACanvas.Pen.Color := clSilver;
        ACanvas.Pen.Style := psDot;
        ACanvas.MoveTo(CX, CY + CH);
        ACanvas.LineTo(CX + CW, CY + CH);
      End;
    End;
  End;

  // Draw selected cell highlight
  For i := 0 To Length(FSelectedCells) - 1 Do
  Begin
    If (FSelectedCells[i].X >= 0) AND (FSelectedCells[i].X < FCols) AND (FSelectedCells[i].Y >= 0) AND (FSelectedCells[i].Y < FRows) Then
    Begin
      ACanvas.Pen.Style := psSolid;
      ACanvas.Pen.Color := clBlue;
      ACanvas.Pen.Width := 2;
      ACanvas.Brush.Style := bsClear;

      IsMerged := IsCellMerged(FSelectedCells[i].Y, FSelectedCells[i].X, MergeInfo);
      If IsMerged Then
      Begin
        X := FLeft;
        For j := 0 To MergeInfo.Col - 1 Do X := X + FColWidths[j];
        Y := FTop;
        For j := 0 To MergeInfo.Row - 1 Do Y := Y + FRowHeights[j];
        CW := 0;
        For j := MergeInfo.Col To MergeInfo.Col + MergeInfo.ColSpan - 1 Do CW := CW + FColWidths[j];
        CH := 0;
        For j := MergeInfo.Row To MergeInfo.Row + MergeInfo.RowSpan - 1 Do CH := CH + FRowHeights[j];
      End
      Else
      Begin
        X := FLeft;
        For j := 0 To FSelectedCells[i].X - 1 Do X := X + FColWidths[j];
        Y := FTop;
        For j := 0 To FSelectedCells[i].Y - 1 Do Y := Y + FRowHeights[j];
        CW := FColWidths[FSelectedCells[i].X];
        CH := FRowHeights[FSelectedCells[i].Y];
      End;

      ACanvas.Rectangle(X, Y, X + CW, Y + CH);
      ACanvas.Pen.Width := 1;
    End;
  End;

  // Draw grid selection
  If FSelected Then
  Begin
    ACanvas.Pen.Style := psDot;
    ACanvas.Pen.Color := clRed;
    ACanvas.Brush.Style := bsClear;
    ACanvas.Rectangle(FLeft, FTop, FLeft + FWidth, FTop + FHeight);
  End;
End;

Function TLabelGrid.Clone: TLabelItem;
Var
  NewGrid: TLabelGrid;
  r, c, i: Integer;
  Child, NewChild: TLabelItem;
Begin
  NewGrid := TLabelGrid.Create;
  NewGrid.Name := Name;
  NewGrid.Left := Left;
  NewGrid.Top := Top;
  NewGrid.Width := Width;
  NewGrid.Height := Height;
  NewGrid.Rows := Rows;
  NewGrid.Cols := Cols;
  NewGrid.Popup := Popup;
  NewGrid.FieldName := FieldName;
  NewGrid.RecordOffset := RecordOffset;

  // Copy cell contents
  For r := 0 To Rows - 1 Do
  Begin
    For c := 0 To Cols - 1 Do
    Begin
      Child := GetCellItem(r, c);
      If Assigned(Child) Then
      Begin
        NewChild := Child.Clone;
        NewGrid.SetCellItem(r, c, NewChild);
      End;

      // Copy cell formatting
      NewGrid.FCellFormats[r, c] := FCellFormats[r, c];
    End;
  End;

  // Copy merged cells
  SetLength(NewGrid.FMergedCells, Length(FMergedCells));
  For i := 0 To Length(FMergedCells) - 1 Do
    NewGrid.FMergedCells[i] := FMergedCells[i];

  // Copy dimensions
  SetLength(NewGrid.FColWidths, FCols);
  For i := 0 To FCols - 1 Do NewGrid.FColWidths[i] := FColWidths[i];
  SetLength(NewGrid.FRowHeights, FRows);
  For i := 0 To FRows - 1 Do NewGrid.FRowHeights[i] := FRowHeights[i];

  Result := NewGrid;
End;

Function TLabelGrid.GetItemType: TLabelItemType;
Begin
  Result := litGrid;
End;

Procedure TLabelGrid.SaveToWriter(AWriter: TWriter);
Var
  r, c: Integer;
  LItem: TLabelItem;
Begin
  AWriter.WriteInteger(Ord(GetItemType));
  Inherited SaveToWriter(AWriter);
  AWriter.WriteInteger(FRows);
  AWriter.WriteInteger(FCols);

  // Cells
  For r := 0 To FRows - 1 Do
    For c := 0 To FCols - 1 Do
    Begin
      LItem := FCells[r, c];
      If Assigned(LItem) Then
      Begin
        AWriter.WriteBoolean(True);
        LItem.SaveToWriter(AWriter);
      End
      Else
        AWriter.WriteBoolean(False);
    End;

  // Formats
  For r := 0 To FRows - 1 Do
    For c := 0 To FCols - 1 Do
    Begin
      AWriter.WriteInteger(FCellFormats[r, c].BackColor);
      AWriter.WriteInteger(FCellFormats[r, c].BorderColor);
      AWriter.WriteInteger(Byte(FCellFormats[r, c].BorderVisible));
    End;

  // Merges
  AWriter.WriteInteger(Length(FMergedCells));
  For r := 0 To High(FMergedCells) Do
  Begin
    AWriter.WriteInteger(FMergedCells[r].Row);
    AWriter.WriteInteger(FMergedCells[r].Col);
    AWriter.WriteInteger(FMergedCells[r].RowSpan);
    AWriter.WriteInteger(FMergedCells[r].ColSpan);
  End;

  // Dimensions
  For r := 0 To FCols - 1 Do AWriter.WriteInteger(FColWidths[r]);
  For r := 0 To FRows - 1 Do AWriter.WriteInteger(FRowHeights[r]);
End;

Procedure TLabelGrid.LoadFromReader(AReader: TReader);
Var
  r, c, LCount: Integer;
  LHasItem: Boolean;
Begin
  Inherited LoadFromReader(AReader);
  FRows := AReader.ReadInteger;
  FCols := AReader.ReadInteger;
  UpdateCells;
  UpdateFormats;

  // Cells
  For r := 0 To FRows - 1 Do
    For c := 0 To FCols - 1 Do
    Begin
      LHasItem := AReader.ReadBoolean;
      If LHasItem Then
        FCells[r, c] := TLabelItem.CreateFromReader(AReader)
      Else
        FCells[r, c] := nil;
    End;

  // Formats
  For r := 0 To FRows - 1 Do
    For c := 0 To FCols - 1 Do
    Begin
      FCellFormats[r, c].BackColor := TColor(AReader.ReadInteger);
      FCellFormats[r, c].BorderColor := TColor(AReader.ReadInteger);
      FCellFormats[r, c].BorderVisible := TBorderEdges(Byte(AReader.ReadInteger));
    End;

  // Merges
  LCount := AReader.ReadInteger;
  SetLength(FMergedCells, LCount);
  For r := 0 To LCount - 1 Do
  Begin
    FMergedCells[r].Row := AReader.ReadInteger;
    FMergedCells[r].Col := AReader.ReadInteger;
    FMergedCells[r].RowSpan := AReader.ReadInteger;
    FMergedCells[r].ColSpan := AReader.ReadInteger;
  End;

  // Dimensions
  SetLength(FColWidths, FCols);
  For r := 0 To FCols - 1 Do FColWidths[r] := AReader.ReadInteger;
  SetLength(FRowHeights, FRows);
  For r := 0 To FRows - 1 Do FRowHeights[r] := AReader.ReadInteger;
End;

{ TLabelManager }

Constructor TLabelManager.Create;
Begin
  FItems := TObjectList.Create(True);
End;

Destructor TLabelManager.Destroy;
Begin
  FItems.Free;
  Inherited Destroy;
End;

Procedure TLabelManager.Clear;
Begin
  FItems.Clear;
End;

Procedure TLabelManager.AddItem(AItem: TLabelItem);
Begin
  FItems.Add(AItem);
End;

Function TLabelManager.GetItem(AIndex: Integer): TLabelItem;
Begin
  Result := TLabelItem(FItems[AIndex]);
End;

Function TLabelManager.GetCount: Integer;
Begin
  Result := FItems.Count;
End;

Function TLabelManager.AddText(AText: String): TLabelText;
Begin
  Result := TLabelText.Create;
  Result.Text := AText;
  FItems.Add(Result);
End;

Function TLabelManager.AddImage: TLabelImage;
Begin
  Result := TLabelImage.Create;
  FItems.Add(Result);
End;

Function TLabelManager.AddBarcode(AText: String; AKind: TBarcodeKind): TLabelBarcode;
Begin
  Result := TLabelBarcode.Create;
  Result.Text := AText;
  Result.Kind := AKind;
  FItems.Add(Result);
End;

Function TLabelManager.AddGrid(ARows, ACols: Integer): TLabelGrid;
Begin
  Result := TLabelGrid.Create;
  Result.Rows := ARows;
  Result.Cols := ACols;
  FItems.Add(Result);
End;

Procedure TLabelManager.RemoveItem(AItem: TLabelItem);
Begin
  FItems.Remove(AItem);
End;

Function TLabelManager.ExistsName(Const AName: String; AExcludeItem: TLabelItem = nil): Boolean;

  Function CheckItem(AItem: TLabelItem): Boolean;
  Var
    r, c: Integer;
    LGrid: TLabelGrid;
    LChild: TLabelItem;
  Begin
    Result := False;
    If (AItem <> AExcludeItem) AND SameText(AItem.Name, AName) Then
    Begin
      Result := True;
      Exit;
    End;

    If AItem IS TLabelGrid Then
    Begin
      LGrid := TLabelGrid(AItem);
      For r := 0 To LGrid.Rows - 1 Do
        For c := 0 To LGrid.Cols - 1 Do
        Begin
          LChild := LGrid.GetCellItem(r, c);
          If Assigned(LChild) Then
            If CheckItem(LChild) Then
            Begin
              Result := True;
              Exit;
            End;
        End;
    End;
  End;

Var
  i: Integer;
Begin
  Result := False;
  For i := 0 To FItems.Count - 1 Do
    If CheckItem(TLabelItem(FItems[i])) Then
    Begin
      Result := True;
      Exit;
    End;
End;

Function TLabelManager.GenerateUniqueName(Const APrefix: String): String;
Var
  i: Integer;
Begin
  i := 1;
  Result := APrefix + IntToStr(i);
  While ExistsName(Result) Do
  Begin
    Inc(i);
    Result := APrefix + IntToStr(i);
  End;
End;

Function TLabelManager.ExtractItem(AItem: TLabelItem): TLabelItem;
Begin
  FItems.Extract(AItem);
  Result := AItem;
End;

Procedure TLabelManager.DrawAll(ACanvas: TCanvas);
Var
  i: Integer;
Begin
  For i := 0 To FItems.Count - 1 Do
    TLabelItem(FItems[i]).Draw(ACanvas);
End;

Function TLabelManager.FindItemAt(AX, AY: Integer): TLabelItem;
Var
  i: Integer;
Begin
  Result := nil;
  For i := FItems.Count - 1 Downto 0 Do
  Begin
    If TLabelItem(FItems[i]).HitTest(AX, AY) Then
    Begin
      Result := TLabelItem(FItems[i]);
      Exit;
    End;
  End;
End;


Procedure TLabelManager.DeselectAll;

  Procedure DeselectItem(AItem: TLabelItem);
  Var
    r, c: Integer;
    LGrid: TLabelGrid;
    LChild: TLabelItem;
  Begin
    AItem.Selected := False;
    If AItem IS TLabelGrid Then
    Begin
      LGrid := TLabelGrid(AItem);
      LGrid.ClearSelection; // Clear cell selection too
      For r := 0 To LGrid.Rows - 1 Do
        For c := 0 To LGrid.Cols - 1 Do
        Begin
          LChild := LGrid.GetCellItem(r, c);
          If Assigned(LChild) Then
            DeselectItem(LChild);
        End;
    End;
  End;

Var
  i: Integer;
Begin
  For i := 0 To FItems.Count - 1 Do
    DeselectItem(TLabelItem(FItems[i]));
End;

Function TLabelManager.FindGridAt(AX, AY: Integer; AExclude: TLabelItem): TLabelGrid;
Var
  i: Integer;
  Item: TLabelItem;
Begin
  Result := nil;
  For i := FItems.Count - 1 Downto 0 Do
  Begin
    Item := TLabelItem(FItems[i]);
    If (Item <> AExclude) AND (Item IS TLabelGrid) AND Item.HitTest(AX, AY) Then
    Begin
      Result := TLabelGrid(Item);
      Exit;
    End;
  End;
End;

Procedure TLabelManager.BringToFront(AItem: TLabelItem);
Var
  Index: Integer;
Begin
  Index := FItems.IndexOf(AItem);
  If Index >= 0 Then
    FItems.Move(Index, FItems.Count - 1);
End;

Procedure TLabelManager.SendToBack(AItem: TLabelItem);
Var
  Index: Integer;
Begin
  Index := FItems.IndexOf(AItem);
  If Index >= 0 Then
    FItems.Move(Index, 0);
End;

Function TLabelManager.DuplicateItem(AItem: TLabelItem): TLabelItem;
Begin
  Result := AItem.Clone;
  Result.Name := GenerateUniqueName(AItem.Name + '_Copy');
  FItems.Add(Result);
End;

Function TLabelManager.GetSelectedItems: TList;

  Procedure CollectSelected(AItem: TLabelItem; AList: TList);
  Var
    r, c: Integer;
    LGrid: TLabelGrid;
    LChild: TLabelItem;
  Begin
    If AItem.Selected Then
      AList.Add(AItem);

    If AItem IS TLabelGrid Then
    Begin
      LGrid := TLabelGrid(AItem);
      For r := 0 To LGrid.Rows - 1 Do
        For c := 0 To LGrid.Cols - 1 Do
        Begin
          LChild := LGrid.GetCellItem(r, c);
          If Assigned(LChild) Then
            CollectSelected(LChild, AList);
        End;
    End;
  End;

Var
  i: Integer;
Begin
  Result := TList.Create;
  For i := 0 To FItems.Count - 1 Do
    CollectSelected(TLabelItem(FItems[i]), Result);
End;

Function TLabelGrid.HitTestGridLine(AX, AY: Integer; Var AIsRow: Boolean; Var AIndex: Integer): Boolean;
Const
  Tolerance = 3;
Var
  i, Pos: Integer;
Begin
  Result := False;
  AIsRow := False;
  AIndex := -1;

  If NOT HitTest(AX, AY) Then Exit;

  // Check vertical lines (columns)
  Pos := FLeft;
  For i := 0 To FCols - 2 Do
  Begin
    Pos := Pos + FColWidths[i];
    If Abs(AX - Pos) <= Tolerance Then
    Begin
      AIsRow := False;
      AIndex := i;
      Result := True;
      Exit;
    End;
  End;

  // Check horizontal lines (rows)
  Pos := FTop;
  For i := 0 To FRows - 2 Do
  Begin
    Pos := Pos + FRowHeights[i];
    If Abs(AY - Pos) <= Tolerance Then
    Begin
      AIsRow := True;
      AIndex := i;
      Result := True;
      Exit;
    End;
  End;
End;

Function TLabelGrid.GetCellAt(AX, AY: Integer; Out ARow, ACol: Integer): Boolean;
Var
  i, LPos: Integer;
Begin
  Result := False;
  ARow := -1;
  ACol := -1;

  If NOT HitTest(AX, AY) Then Exit;

  // Find Column
  LPos := FLeft;
  For i := 0 To FCols - 1 Do
  Begin
    If (AX >= LPos) AND (AX < LPos + FColWidths[i]) Then
    Begin
      ACol := i;
      Break;
    End;
    LPos := LPos + FColWidths[i];
  End;

  // Find Row
  LPos := FTop;
  For i := 0 To FRows - 1 Do
  Begin
    If (AY >= LPos) AND (AY < LPos + FRowHeights[i]) Then
    Begin
      ARow := i;
      Break;
    End;
    LPos := LPos + FRowHeights[i];
  End;

  Result := (ARow <> -1) AND (ACol <> -1);
End;


Procedure TLabelGrid.SetColWidth(AIndex, AWidth: Integer);
Begin
  If (AIndex >= 0) AND (AIndex < FCols) Then
    FColWidths[AIndex] := AWidth;
End;

Procedure TLabelGrid.SetRowHeight(AIndex, AHeight: Integer);
Begin
  If (AIndex >= 0) AND (AIndex < FRows) Then
    FRowHeights[AIndex] := AHeight;
End;

Function TLabelGrid.GetColWidth(AIndex: Integer): Integer;
Begin
  If (AIndex >= 0) AND (AIndex < FCols) Then
    Result := FColWidths[AIndex]
  Else
    Result := 0;
End;

Function TLabelGrid.GetRowHeight(AIndex: Integer): Integer;
Begin
  If (AIndex >= 0) AND (AIndex < FRows) Then
    Result := FRowHeights[AIndex]
  Else
    Result := 0;
End;

End.
