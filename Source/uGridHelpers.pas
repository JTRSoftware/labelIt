Unit uGridHelpers;

{$mode Delphi}{$H+}

Interface

Uses
  Classes, SysUtils, Grids, Math;

Type
  { TStringGridHelper }
  TStringGridHelper = Class helper For TStringGrid
  Public
    { Garante que a soma das colunas não ultrapassa a largura visível da grid }
    Procedure AjustarLarguraAoLimite;
    Function GetTotalColWidths: Integer;
  End;

Implementation

{ TStringGridHelper }

Procedure TStringGridHelper.AjustarLarguraAoLimite;
Var
  i, LTotalWidth, LGridWidth, LCurrentSum, LNewWidth: Integer;
Begin
  If ColCount <= 0 Then Exit;

  // LCL ClientWidth já exclui o espaço das scrollbars visíveis
  LGridWidth := Self.ClientWidth - 2; // Margem de segurança de 2 pixels
  
  LTotalWidth := GetTotalColWidths;

  // Só intervimos se as colunas estiverem a transbordar
  If (LTotalWidth > LGridWidth) AND (LTotalWidth > 0) Then
  Begin
    Self.BeginUpdate;
    Try
      LCurrentSum := 0;
      For i := 0 To ColCount - 1 Do
      Begin
        // Calculamos a nova largura proporcional
        LNewWidth := Round((Int64(ColWidths[i]) * LGridWidth) / LTotalWidth);
        
        // Garantir um mínimo para não desaparecerem colunas
        If LNewWidth < 10 Then LNewWidth := 10;
        
        // Na última coluna, fazemos o ajuste fino para bater certo ao pixel
        If i = ColCount - 1 Then
          ColWidths[i] := Max(10, LGridWidth - LCurrentSum)
        Else
        Begin
          ColWidths[i] := LNewWidth;
          LCurrentSum := LCurrentSum + LNewWidth;
        End;
      End;
    Finally
      Self.EndUpdate;
    End;
  End;
End;

Function TStringGridHelper.GetTotalColWidths: Integer;
Var
  i: Integer;
Begin
  Result := 0;
  For i := 0 To ColCount - 1 Do
    Result := Result + ColWidths[i];
End;

End.
