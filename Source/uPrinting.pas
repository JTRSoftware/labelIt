Unit uPrinting;

{$mode Delphi}

Interface

Uses
  Classes, SysUtils, Graphics, Printers, uLabelObjects;

Type
  { TLabelPrinter }

  TLabelPrinter = Class
  Public
    Class Procedure PrintLabel(AManager: TLabelManager; APrinterName: String = '');
  End;

Implementation

{ TLabelPrinter }

Class Procedure TLabelPrinter.PrintLabel(AManager: TLabelManager; APrinterName: String);
Begin
  If APrinterName <> '' Then
    Printer.SetPrinter(APrinterName);

  Printer.BeginDoc;
  Try
    AManager.DrawAll(Printer.Canvas);
  Finally
    Printer.EndDoc;
  End;
End;

End.
