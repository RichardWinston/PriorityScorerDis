unit frmMainUnit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Grids,
  RbwDataGrid4, System.Generics.Collections, Vcl.ComCtrls, Vcl.Mask, JvExMask,
  JvToolEdit;

type
  TScoringMatrix = array of array of integer;
  TfrmMain = class(TForm)
    pnl1: TPanel;
    btn1: TButton;
    ed1: TLabeledEdit;
    pc1: TPageControl;
    tabScoring_Matrix: TTabSheet;
    tabPriorities: TTabSheet;
    rdgPriorities: TRbwDataGrid4;
    rdgScoringMatrix: TRbwDataGrid4;
    btnAnalyze: TButton;
    fedExcel: TJvFilenameEdit;
    lbl1: TLabel;
    fedCSV: TJvFilenameEdit;
    lbl2: TLabel;
    procedure btn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure rdgPrioritiesBeforeDrawCell(Sender: TObject; ACol, ARow: Integer);
    procedure btnAnalyzeClick(Sender: TObject);
  private
    procedure CalculateData(AXLSFile: string);
    procedure InitializeScoringMatrix(var ScoringMatrix: TScoringMatrix);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

const
  TruePriorities: array[0..9] of Integer = (1,1,2,2,2,2,3,3,3,3);
  xlWorkbookDefault = 51;

implementation

uses
  Vcl.OleAuto;

{$R *.dfm}

procedure TfrmMain.btn1Click(Sender: TObject);
var
  AnInt: Integer;
  RowIndex: Integer;
  Tier: Integer;
  Value: Integer;
  Score: Integer;
  IntList: TList<Integer>;
  ScoringMatrix: TScoringMatrix;
begin
  Score := 0;
  InitializeScoringMatrix(ScoringMatrix);
  IntList := TList<Integer>.Create;
  try
    for RowIndex := 1 to rdgPriorities.RowCount - 1 do
    begin
      AnInt := StrToIntDef(rdgPriorities.Cells[1,RowIndex], 11);
      if AnInt = 0 then
      begin
        AnInt := 11;
      end;
      if AnInt <> 11 then
      begin
        Assert(IntList.IndexOf(AnInt) < 0, 'No two items should receive the same rank.');
        IntList.Add(AnInt);
      end;
      case RowIndex of
        1..2: Tier := 1;
        3..6: Tier := 2;
        7..10: Tier := 3;
        else Assert(False);
      end;
      Value := ScoringMatrix[AnInt-1, Tier-1];
      Score := Score+ Value;
    end;
  finally
    IntList.Free;
  end;
  ed1.Text := IntToStr(Score);
end;

procedure TfrmMain.btnAnalyzeClick(Sender: TObject);
begin
  if FileExists(fedExcel.FileName) and (fedCSV.FileName <> '') then
  begin
    CalculateData(fedExcel.FileName);
    ShowMessage('Done');
  end
  else
  begin
    Beep;
  end;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  index: Integer;
begin
  pc1.ActivePageIndex := 0;

  rdgScoringMatrix.BeginUpdate;
  try
    for index := 1 to 10 do
    begin
      rdgScoringMatrix.Cells[index,0] := IntToStr(index);
    end;
    rdgScoringMatrix.Cells[11,0] := 'Missing or Wrong';

    rdgScoringMatrix.Cells[0,1] := 'Tier 1';
    rdgScoringMatrix.Cells[1,1] := '10';
    rdgScoringMatrix.Cells[2,1] := '10';
    rdgScoringMatrix.Cells[3,1] := '9';
    rdgScoringMatrix.Cells[4,1] := '8';
    rdgScoringMatrix.Cells[5,1] := '7';
    rdgScoringMatrix.Cells[6,1] := '6';
    rdgScoringMatrix.Cells[7,1] := '5';
    rdgScoringMatrix.Cells[8,1] := '4';
    rdgScoringMatrix.Cells[9,1] := '3';
    rdgScoringMatrix.Cells[10,1] := '2';
    rdgScoringMatrix.Cells[11,1] := '-5';

    rdgScoringMatrix.Cells[0,2] := 'Tier 2';
    rdgScoringMatrix.Cells[1,2] := '3';
    rdgScoringMatrix.Cells[2,2] := '4';
    rdgScoringMatrix.Cells[3,2] := '5';
    rdgScoringMatrix.Cells[4,2] := '5';
    rdgScoringMatrix.Cells[5,2] := '5';
    rdgScoringMatrix.Cells[6,2] := '5';
    rdgScoringMatrix.Cells[7,2] := '4';
    rdgScoringMatrix.Cells[8,2] := '3';
    rdgScoringMatrix.Cells[9,2] := '2';
    rdgScoringMatrix.Cells[10,2] := '1';
    rdgScoringMatrix.Cells[11,2] := '-3';

    rdgScoringMatrix.Cells[0,3] := 'Tier 3';
    rdgScoringMatrix.Cells[1,3] := '-3';
    rdgScoringMatrix.Cells[2,3] := '-2';
    rdgScoringMatrix.Cells[3,3] := '-1';
    rdgScoringMatrix.Cells[4,3] := '0';
    rdgScoringMatrix.Cells[5,3] := '1';
    rdgScoringMatrix.Cells[6,3] := '2';
    rdgScoringMatrix.Cells[7,3] := '3';
    rdgScoringMatrix.Cells[8,3] := '3';
    rdgScoringMatrix.Cells[9,3] := '3';
    rdgScoringMatrix.Cells[10,3] := '3';
    rdgScoringMatrix.Cells[11,3] := '-1';
  finally
    rdgScoringMatrix.EndUpdate;
  end;

  rdgPriorities.BeginUpdate;
  try
    rdgPriorities.Cells[0,1] := 'Fetal Distress';
    rdgPriorities.Cells[0,2] := 'Seizure/ PreE';
    rdgPriorities.Cells[0,3] := 'Breech';
    rdgPriorities.Cells[0,4] := 'Cesarean';
    rdgPriorities.Cells[0,5] := 'Pain';
    rdgPriorities.Cells[0,6] := 'Preterm/ GA';
    rdgPriorities.Cells[0,7] := 'PCN/ Risk';
    rdgPriorities.Cells[0,8] := 'Labor Status';
    rdgPriorities.Cells[0,9] := 'Teen Support System Single';
    rdgPriorities.Cells[0,10] := 'Primip Knowl Deficit';

    rdgPriorities.Cells[1,0] := 'Assigned Priority';
  finally
    rdgPriorities.EndUpdate;
  end;
end;

procedure TfrmMain.rdgPrioritiesBeforeDrawCell(Sender: TObject; ACol,
  ARow: Integer);
begin
  if (ACol = 1) and (ARow >= 1) and (rdgPriorities.Cells[ACol,ARow] <> '')
    and (rdgPriorities.Cells[ACol,ARow] <> '0')
    and (rdgPriorities.Cols[1].IndexOf(rdgPriorities.Cells[ACol,ARow]) < ARow) then
  begin
    rdgPriorities.Canvas.Brush.Color := clRed;
  end;
end;

procedure TfrmMain.CalculateData(AXLSFile: string);
const
  xlCellTypeLastCell = $0000000B;
  ProblemLabelCol = 1;
  ProblemPresentCol = 2;
  ProblemRepeatedCol = 3;
  WrongProblemCol = 4;
  AssignedPriorityCol = 6;
  PriorityScoreCol = 7;
  EvidenceKeyCol = 8;
  EvidenceCol = 9;
  InterventionsKeyCol = 10;
  InterventionsCol = 11;
  OutcomesKeyCol = 12;
  OutcomesCol = 13;

  FirstDataRowPartA = 5;
  LastDataRowPartA = 14;
  TotalsPartARow = 15;
  PercentPartARow = 16;
  FirstDataRowPartB = 19;
  LastDataRowPartB = 21;
  TotalsPartBRow = 22;
  IDRow = 1;
  IDCol = 14;
  RaterIDCol = 9;
  PartBValueColKey = 2;
  PartBValueCol = 3;
  PartBValuePercentCol = 4;

  TotalRow = 22;
  TotalCol = 17;

  MinimumPriorityScore = -5*2 + -3*4 -1*2 -3 -2 -1*8;
  MaximumPriorityScore = 10*2 + 5*4 + 3*4;
  PriorityScoreRange = MaximumPriorityScore-MinimumPriorityScore;

  ColumnLabels: array[0..9] of string = ('Fetal_Distress', 'PreE',
    'Breech', 'Cesarean', 'Pain', 'Preterm/GA', 'PNC/ Risk', 'Labor_Status',
    'Teen_Support_System_Single', 'Primip_Knowl_Deficit');
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  x, y: Integer;
  RowIndex: Integer;
  ColIndex: Integer;
  ParticipantID: string;
  SheetCount: Integer;
  SheetIndex: integer;
  LineBuilder: TStringBuilder;
  OutputFile: TStringList;
  AString: string;
  TotalProblems: Integer;
  TotalRepeated: Integer;
  Present: Integer;
  PercentProblems: double;
  PercentWrongProbs: double;
  WrongProbs: Integer;
  Priority: Integer;
  Tier: Integer;
  ScoringMatrix: TScoringMatrix;
  PriorityScore: integer;
  PriorityScoreTotal: Integer;
  RaterID: string;
  Repeated: Integer;
  PriorityScorePercent: double;
  Evidence: Integer;
  TotalEvidence: Integer;
  TotalEvidenceKey: Integer;
  TotalEvidencePercent: double;
  TotalInterventionsKey: Integer;
  Interventions: Integer;
  TotalOutcomesKey: Integer;
  Outcomes: Integer;
  TotalInterventions: Integer;
  Intervetions: Integer;
  TotalInterventionsPercent: double;
  TotalOutcomes: Integer;
  TotalOutcomesPercent: double;
  PartBKey: array of integer;
  PartBValues: array of integer;
  PartBKeyValue: Integer;
  PartBKeyTotal: Integer;
  AValue: Integer;
  APercent: double;
  Total: Double;
  TitleLine: string;
begin
  SetLength(PartBKey, LastDataRowPartB-FirstDataRowPartB+1);
  SetLength(PartBValues, LastDataRowPartB-FirstDataRowPartB+1);
  InitializeScoringMatrix(ScoringMatrix);
  XLApp := CreateOleObject('Excel.Application');
  LineBuilder := TStringBuilder.Create;
  OutputFile := TStringList.Create;
  try
    XLApp.Visible := False;
    XLApp.Workbooks.Open(AXLSFile);
    Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];
    Sheet.Select;
    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    x := XLApp.ActiveCell.Row;
    y := XLApp.ActiveCell.Column;

    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;

    TotalEvidenceKey := 0;
    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := RangeMatrix[RowIndex,EvidenceKeyCol];
      if AString = '' then
      begin
        AString := '0';
      end;
      Evidence := StrToInt(AString);
      TotalEvidenceKey := TotalEvidenceKey + Evidence;
    end;

    TotalInterventionsKey := 0;
    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := RangeMatrix[RowIndex,InterventionsKeyCol];
      if AString = '' then
      begin
        AString := '0';
      end;
      Interventions := StrToInt(AString);
      TotalInterventionsKey := TotalInterventionsKey + Interventions;
    end;

    TotalOutcomesKey := 0;
    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := RangeMatrix[RowIndex,OutcomesKeyCol];
      if AString = '' then
      begin
        AString := '0';
      end;
      Outcomes := StrToInt(AString);
      TotalOutcomesKey := TotalOutcomesKey + Outcomes;
    end;

    PartBKeyTotal := 0;
    for RowIndex := FirstDataRowPartB to LastDataRowPartB do
    begin
      AString := RangeMatrix[RowIndex,PartBValueColKey];
      PartBKeyValue := StrToInt(AString);
      PartBKey[RowIndex-FirstDataRowPartB] := PartBKeyValue;
      PartBKeyTotal := PartBKeyTotal + PartBKeyValue;
    end;

    LineBuilder.Append('StudyID, ');
    LineBuilder.Append('Rater, ');

    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := ColumnLabels[RowIndex-FirstDataRowPartA];
      LineBuilder.Append(AString);
      LineBuilder.Append('_Present, ');
    end;

    LineBuilder.Append('Total_Problems_Identified, ');
    LineBuilder.Append('Problems_Identified%, ');

    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := ColumnLabels[RowIndex-FirstDataRowPartA];
//      AString := RangeMatrix[RowIndex,ProblemLabelCol];
      LineBuilder.Append(AString);
      LineBuilder.Append('_Repeated, ');
    end;

    LineBuilder.Append('Total_Problems_Repeated, ');

    LineBuilder.Append('NumWrong, ');
    LineBuilder.Append('NumWrong%, ');

    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := ColumnLabels[RowIndex-FirstDataRowPartA];
      LineBuilder.Append(AString);
      LineBuilder.Append('_Priority_Assigned, ');
    end;

    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := ColumnLabels[RowIndex-FirstDataRowPartA];
      LineBuilder.Append(AString);
      LineBuilder.Append('_Priority_Score, ');
    end;

    LineBuilder.Append('Total_Priority_Score, ');

    LineBuilder.Append('Total_Priority_Score_%, ');

    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := ColumnLabels[RowIndex-FirstDataRowPartA];
      LineBuilder.Append(AString);
      LineBuilder.Append('_Evidence, ');
    end;
    LineBuilder.Append('"Total_Evidence", ');
    LineBuilder.Append('"Evidence%", ');

    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := ColumnLabels[RowIndex-FirstDataRowPartA];
      LineBuilder.Append(AString);
      LineBuilder.Append('_Interventions, ');
    end;

    LineBuilder.Append('Interventions_Total, ');
    LineBuilder.Append('Interventions%, ');

    for RowIndex := FirstDataRowPartA to LastDataRowPartA do
    begin
      AString := ColumnLabels[RowIndex-FirstDataRowPartA];
      LineBuilder.Append(AString);
      LineBuilder.Append('_Outcomes_ID''d, ');
    end;
    LineBuilder.Append('Outcomes_Total, ');
    LineBuilder.Append('Outcomes%, ');

    for RowIndex := FirstDataRowPartB to LastDataRowPartB do
    begin
      AString := RangeMatrix[RowIndex,ProblemLabelCol];
      AString := StringReplace(AString, ' ', '_', [rfReplaceAll, rfIgnoreCase]);
      LineBuilder.Append(AString);
      LineBuilder.Append('_ID''d, ');
    end;
    for RowIndex := FirstDataRowPartB to LastDataRowPartB do
    begin
      AString := RangeMatrix[RowIndex,ProblemLabelCol];
      AString := StringReplace(AString, ' ', '_', [rfReplaceAll, rfIgnoreCase]);
      LineBuilder.Append(AString);
      LineBuilder.Append('_ID''d%, ');
    end;

    LineBuilder.Append('Total');

    TitleLine := LineBuilder.ToString;
//    OutputFile.Add();
    LineBuilder.Clear;


    SheetCount := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets.Count;
    for SheetIndex := 1 to SheetCount do
    begin
      Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[SheetIndex];
      Sheet.Select;
      Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;

      RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;

      ParticipantID := RangeMatrix[IDRow,IDCol];
      LineBuilder.Append(ParticipantID);
      LineBuilder.Append(', ');

      RaterID := RangeMatrix[IDRow,RaterIDCol];
      LineBuilder.Append(RaterID);
      LineBuilder.Append(', ');

      TotalProblems := 0;
      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := RangeMatrix[RowIndex,ProblemPresentCol];
        if AString = '' then
        begin
          AString := '0';
        end;
        LineBuilder.Append(AString);
        LineBuilder.Append(', ');
        Present := StrToInt(AString);
        if not (Present in [0,1]) then
        begin
          Beep;
          MessageDlg(Format(
            'Incorrect coding for "problem present" for %s', [ParticipantID]),
            mtError, [mbOK], 0);
          Exit;
        end;
        TotalProblems := TotalProblems + Present;
      end;
  //  Total Problems Identified
      LineBuilder.Append(TotalProblems);
      LineBuilder.Append(', ');
      Sheet.Cells[TotalsPartARow, ProblemPresentCol].Value := TotalProblems;

  //  Percent Problems Identified);
      PercentProblems := TotalProblems/10;
      LineBuilder.Append(PercentProblems);
      LineBuilder.Append(', ');
      Sheet.Cells[PercentPartARow, ProblemPresentCol].NumberFormat := '0%';
      Sheet.Cells[PercentPartARow, ProblemPresentCol].Value := PercentProblems;

      TotalRepeated := 0;
      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := RangeMatrix[RowIndex,ProblemRepeatedCol];
        if AString = '' then
        begin
          AString := '0';
        end;
        LineBuilder.Append(AString);
        LineBuilder.Append(', ');
        Repeated := StrToInt(AString);
        TotalRepeated := TotalRepeated + Repeated;
      end;

//    Total Problems Repeated
      LineBuilder.Append(TotalRepeated);
      LineBuilder.Append(', ');
      Sheet.Cells[TotalsPartARow, ProblemRepeatedCol].Value := TotalRepeated;

//    Unkeyed Score
      AString := UpperCase(RangeMatrix[TotalsPartARow,WrongProblemCol]);
      if AString = '' then
      begin
        AString := '0';
      end;
      WrongProbs := StrToInt(AString);
      LineBuilder.Append(WrongProbs);
      LineBuilder.Append(', ');
      Sheet.Cells[TotalsPartARow, WrongProblemCol].Value := WrongProbs;

//    Unkeyed Score Percent
      PercentWrongProbs := (10-WrongProbs)/10;
      LineBuilder.Append(PercentWrongProbs);
      LineBuilder.Append(', ');
      Sheet.Cells[PercentPartARow, WrongProblemCol].NumberFormat := '0%';
      Sheet.Cells[PercentPartARow, WrongProblemCol].Value := PercentWrongProbs;

      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := RangeMatrix[RowIndex,AssignedPriorityCol];
        if AString = '' then
        begin
          Priority := 0;
        end
        else
        begin
          Priority := StrToInt(AString);
        end;
        LineBuilder.Append(Priority);
        LineBuilder.Append(', ');
      end;


      PriorityScoreTotal := 0;
      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := RangeMatrix[RowIndex,AssignedPriorityCol];
        if AString = '' then
        begin
          Priority := 11;
        end
        else
        begin
          Priority := StrToInt(AString);
          if Priority = 0 then
          begin
            Priority := 11;
          end;
        end;
        case RowIndex - FirstDataRowPartA + 1 of
          1..2:
            begin
              Tier := 1;
            end;
          3..6:
            begin
              Tier := 2;
            end;
          7..10:
            begin
              Tier := 3;
            end;
          else Assert(False);
        end;
        PriorityScore := ScoringMatrix[Priority-1, Tier-1];
        LineBuilder.Append(PriorityScore);
        LineBuilder.Append(', ');
        PriorityScoreTotal := PriorityScoreTotal + PriorityScore;
        Sheet.Cells[RowIndex, PriorityScoreCol].Value := PriorityScore;
      end;

      PriorityScoreTotal := PriorityScoreTotal - WrongProbs;
      LineBuilder.Append(PriorityScoreTotal);
      LineBuilder.Append(', ');
      Sheet.Cells[TotalsPartARow, PriorityScoreCol].Value := PriorityScoreTotal;

      PriorityScorePercent := (PriorityScoreTotal-MinimumPriorityScore)
        /PriorityScoreRange;
      LineBuilder.Append(PriorityScorePercent);
      LineBuilder.Append(', ');
      Sheet.Cells[PercentPartARow, PriorityScoreCol].NumberFormat := '0%';
      Sheet.Cells[PercentPartARow, PriorityScoreCol].Value := PriorityScorePercent;

      TotalEvidence := 0;
      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
//        LineBuilder.Append('"');
        AString := RangeMatrix[RowIndex,EvidenceCol];
        if AString = '' then
        begin
          AString := '0';
        end;
        LineBuilder.Append(AString);
        LineBuilder.Append(', ');
        Evidence := StrToInt(AString);
        TotalEvidence := TotalEvidence + Evidence;
      end;

//    Total Evidence
      LineBuilder.Append(TotalEvidence);
      LineBuilder.Append(', ');
      Sheet.Cells[TotalsPartARow, EvidenceCol].Value := TotalEvidence;

//    Evidence Percent
      TotalEvidencePercent := TotalEvidence/TotalEvidenceKey;
      LineBuilder.Append(TotalEvidencePercent);
      LineBuilder.Append(', ');
      Sheet.Cells[PercentPartARow, EvidenceCol].NumberFormat := '0%';
      Sheet.Cells[PercentPartARow, EvidenceCol].Value := TotalEvidencePercent;

      TotalInterventions := 0;
      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
//        LineBuilder.Append('"');
        AString := RangeMatrix[RowIndex,InterventionsCol];
        if AString = '' then
        begin
          AString := '0';
        end;
        LineBuilder.Append(AString);
        LineBuilder.Append(', ');
        Intervetions := StrToInt(AString);
        TotalInterventions := TotalInterventions + Intervetions;
      end;

      LineBuilder.Append(TotalInterventions);
      LineBuilder.Append(', ');
      Sheet.Cells[TotalsPartARow, InterventionsCol].Value := TotalInterventions;

      TotalInterventionsPercent := TotalInterventions/TotalInterventionsKey;
      LineBuilder.Append(TotalInterventionsPercent);
      LineBuilder.Append(', ');
      Sheet.Cells[PercentPartARow, InterventionsCol].NumberFormat := '0%';
      Sheet.Cells[PercentPartARow, InterventionsCol].Value := TotalInterventionsPercent;

      TotalOutcomes := 0;
      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
//        LineBuilder.Append('"');
        AString := RangeMatrix[RowIndex,OutcomesCol];
        if AString = '' then
        begin
          AString := '0';
        end;
        LineBuilder.Append(AString);
        LineBuilder.Append(', ');
        Outcomes := StrToInt(AString);
        TotalOutcomes := TotalOutcomes + Outcomes;
      end;
      Sheet.Cells[TotalsPartARow, OutcomesCol].Value := TotalOutcomes;

      LineBuilder.Append(TotalOutcomes);
      LineBuilder.Append(', ');

      TotalOutcomesPercent := TotalOutcomes/TotalOutcomesKey;
      LineBuilder.Append(TotalOutcomesPercent);
      LineBuilder.Append(', ');
      Sheet.Cells[PercentPartARow, OutcomesCol].NumberFormat := '0%';
      Sheet.Cells[PercentPartARow, OutcomesCol].Value := TotalOutcomesPercent;

      for RowIndex := FirstDataRowPartB to LastDataRowPartB do
      begin
        AString := RangeMatrix[RowIndex,PartBValueCol];
        if AString = '' then
        begin
          AString := '0';
        end;
        LineBuilder.Append(AString);
        LineBuilder.Append(', ');
        AValue := StrToInt(AString);
        PartBValues[RowIndex-FirstDataRowPartB] := AValue;
      end;

      Total := PercentProblems + PercentWrongProbs + PriorityScorePercent
        + TotalEvidencePercent + TotalInterventionsPercent + TotalOutcomesPercent;
      for RowIndex := FirstDataRowPartB to LastDataRowPartB do
      begin
        APercent := PartBValues[RowIndex-FirstDataRowPartB]/
          PartBKey[RowIndex-FirstDataRowPartB];
        LineBuilder.Append(APercent);
        LineBuilder.Append(', ');
        Sheet.Cells[RowIndex, PartBValuePercentCol].NumberFormat := '0%';
        Sheet.Cells[RowIndex, PartBValuePercentCol].Value := APercent;

        Total := Total + APercent;
      end;
      Total := Total / 9;
      LineBuilder.Append(Total);
      Sheet.Cells[TotalRow, TotalCol].NumberFormat := '0%';
      Sheet.Cells[TotalRow, TotalCol].Value := Total;

      Sheet.Range['A1','A1'].Select;

      RangeMatrix := Unassigned;
      OutputFile.Add(LineBuilder.ToString);
      LineBuilder.Clear;
    end;

    OutputFile.Sort;
    OutputFile.Insert(0, TitleLine);

    Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];
    Sheet.Select;

    OutputFile.SaveToFile(fedCSV.FileName);
//    XLApp.Workbooks[ExtractFileName(AXLSFile)].SaveAs(AXLSFile, xlWorkbookDefault);
    XLApp.DisplayAlerts := False;
    XLApp.Workbooks[ExtractFileName(AXLSFile)].Save;
  finally
    LineBuilder.Free;
    OutputFile.Free;
    if not VarIsEmpty(XLApp) then
    begin
      XLApp.Quit;
      Sheet := Unassigned;
      XLAPP := Unassigned;
    end;
  end;
end;

procedure TfrmMain.InitializeScoringMatrix(var ScoringMatrix: TScoringMatrix);
var
  ColIndex: Integer;
  RowIndex: Integer;
begin
  SetLength(ScoringMatrix, rdgScoringMatrix.ColCount - 1, rdgScoringMatrix.RowCount - 1);
  for RowIndex := 1 to rdgScoringMatrix.RowCount - 1 do
  begin
    for ColIndex := 1 to rdgScoringMatrix.ColCount - 1 do
    begin
      ScoringMatrix[ColIndex - 1, RowIndex - 1] := rdgScoringMatrix.IntegerValue[ColIndex, RowIndex];
    end;
  end;
end;

end.
