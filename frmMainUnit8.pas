unit frmMainUnit8;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Grids,
  RbwDataGrid4, System.Generics.Collections, Vcl.ComCtrls, Vcl.Mask, JvExMask,
  JvToolEdit;

type
  TIntArray = array of Integer;
  TDoubleArray = array of double;
  TScoringMatrix = array of array of integer;
  TfrmMain = class(TForm)
    pc1: TPageControl;
    tabScoring_Matrix: TTabSheet;
    tabPriorities: TTabSheet;
    rdgPriorities: TRbwDataGrid4;
    rdgScoringMatrix: TRbwDataGrid4;
    tabWarningsOrErrors: TTabSheet;
    memoErrors: TMemo;
    pbCompletion: TProgressBar;
    pnl1: TPanel;
    lbl1: TLabel;
    lbl2: TLabel;
    lblMax: TLabel;
    lblMin: TLabel;
    lbl3: TLabel;
    btnAnalyze: TButton;
    fedExcel: TJvFilenameEdit;
    fedCSV: TJvFilenameEdit;
    btnCopyScoringMatrix: TButton;
    pnl2: TPanel;
    btn1: TButton;
    procedure btn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure rdgPrioritiesBeforeDrawCell(Sender: TObject; ACol, ARow: Integer);
    procedure btnAnalyzeClick(Sender: TObject);
    procedure btnCopyScoringMatrixClick(Sender: TObject);
  private
    procedure CalculateData(AXLSFile: string);
    procedure InitializeScoringMatrix(var ScoringMatrix: TScoringMatrix);
    procedure StoreEvidenceTypeValues(LineBuilder: TStringBuilder;
      AnArray: TIntArray);
    procedure CountEvidence(LineBuilder: TStringBuilder; AnArray: TIntArray;
      var Evidence: Integer; const ErrorLabel, ParticipantID, ProblemType, SheetName: string);
    procedure InitializeEvidenceArray(var AnArray: TIntArray; MaxKey: integer);
    procedure ReadEvidenceTypes(AnArray: TIntArray; ParticipantID: string;
      AString: string; EvidenceTypes: TStringList;
      const ErrorLabel, ProblemType, SheetName: string);
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
  Vcl.OleAuto, System.Math;

{$R *.dfm}

const
  xlCellTypeLastCell = $0000000B;
  ProblemLabelCol = 1;
  PartBProblemLabelRow = 2;
  ProblemPresentCol = 2;
  ProblemRepeatedCol = 3;
  WrongProblemCol = 4;
  AssignedPriorityCol = 6;
  PriorityScoreCol = 7;

  EvidenceKeyCol = 8;
  EvidenceCol = 9;
  EvidencePercentCol = 10;

  InterventionsKeyCol = 24;  //10;
  InterventionsCol = 12; //11;
  InterventionsPercentCol = 13;

  OutcomesKeyCol = 14; //12;
  OutcomesCol = 15; //13;
  OutcomesPercentCol = 16;

  FirstDataRowPartA = 3;  //5;
  LastDataRowPartA = 12;  // 14;
  TotalsPartARow = 13; // 15;
  AdjustedPriorityPointsRow = 14;
  RepeatWrongProbRow = 15;
  PrioritySubScoreRow = 16;
//  PercentPartARow = 14; // 16;
  PriorityPercentRow = 17;


  RepeatWrongProbCol = 7;

//  FirstDataRowPartB = 19;
//  LastDataRowPartB = 21;

  FirstDataColPartB = 18;
  LastDataColPartB = 20;

//  TotalsPartBRow = 22;
  IDRow = 1;
  IDCol = 16; //14;
  RaterIDCol = 9;
  DateCol = 19;

  PartBLbaelRow = 2;
//  PartBValueColKey = 2;
  PartBValueRowKey = 4;
//  PartBValueCol = 3;
  PartBValueRow = 3;
//  PartBValuePercentCol = 4;
  PartBValuePercentRow = 6;

  TotalRow = 18; //15; //22;
  TotalCol = 20; //17;


  ProblemLabels: array[0..9] of string = ('FetalDistress', 'PreE',
    'Breech', 'Cesarean', 'Pain', 'Preterm', 'PNC', 'LaborStatus',
    'Teen', 'KnowlDeficit');

  PartBLabels: array[0..2] of string = ('LegalEthical', 'Missing', 'DC');

procedure TfrmMain.btn1Click(Sender: TObject);
var
  AnInt: Integer;
  RowIndex: Integer;
  Tier: Integer;
  Value: Integer;
  Score: Integer;
  IntList: TList<Integer>;
  ScoringMatrix: TScoringMatrix;
  ColIndex: Integer;
begin
  InitializeScoringMatrix(ScoringMatrix);
  IntList := TList<Integer>.Create;
  try
    for ColIndex := 1 to rdgPriorities.ColCount -1 do
    begin
      Score := 0;
      IntList.Clear;
      for RowIndex := 1 to rdgPriorities.RowCount - 2 do
      begin
        AnInt := StrToIntDef(rdgPriorities.Cells[ColIndex,RowIndex], 11);
        if RowIndex <> rdgPriorities.RowCount - 2 then
        begin
          if AnInt = 0 then
          begin
            AnInt := 11;
          end;
          if AnInt <> 11 then
          begin
            Assert(IntList.IndexOf(AnInt) < 0, 'No two items should receive the same rank.');
            IntList.Add(AnInt);
          end;
          Tier := -1;

          case RowIndex of
            1..2: Tier := 1;
            3..6: Tier := 2;
            7..10: Tier := 3;
            else Assert(False);
          end;
          Value := ScoringMatrix[AnInt-1, Tier-1];
        end
        else
        begin
          Value := -AnInt;
        end;
        Score := Score+ Value;
      end;
      rdgPriorities.Cells[ColIndex,12] := IntToStr(Score);
    end;
  finally
    IntList.Free;
  end;
end;

procedure TfrmMain.btnAnalyzeClick(Sender: TObject);
begin
  if FileExists(fedExcel.FileName) and (fedCSV.FileName <> '') then
  begin
    memoErrors.Lines.Clear;
    try
      CalculateData(fedExcel.FileName);
    finally
      if memoErrors.Lines.Count > 0 then
      begin
        tabWarningsOrErrors.TabVisible := True;
        pc1.ActivePage := tabWarningsOrErrors;
      end;
      Beep;
      ShowMessage('Done');
    end;
  end
  else
  begin
    Beep;
    ShowMessage('You need to give names for both the input and output files.');
  end;
end;

procedure TfrmMain.btnCopyScoringMatrixClick(Sender: TObject);
begin
  rdgScoringMatrix.CopyAllCellsToClipboard;
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
    rdgScoringMatrix.Cells[11,1] := '0';

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
    rdgScoringMatrix.Cells[11,2] := '0';

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
    rdgScoringMatrix.Cells[11,3] := '0';
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
    rdgPriorities.Cells[0,11] := 'WRONG';
    rdgPriorities.Cells[0,12] := 'Priority Score';

    rdgPriorities.Cells[1,0] := 'Assigned Priority';
  finally
    rdgPriorities.EndUpdate;
  end;
end;

procedure TfrmMain.rdgPrioritiesBeforeDrawCell(Sender: TObject; ACol,
  ARow: Integer);
begin
  if (ACol = 1) and (ARow >= 1) and (ARow < rdgPriorities.RowCount-1)
    and (rdgPriorities.Cells[ACol,ARow] <> '')
    and (rdgPriorities.Cells[ACol,ARow] <> '0')
    and (rdgPriorities.Cols[1].IndexOf(rdgPriorities.Cells[ACol,ARow]) < ARow) then
  begin
    rdgPriorities.Canvas.Brush.Color := clRed;
  end;
end;

procedure TfrmMain.CalculateData(AXLSFile: string);
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
  PartBKey: TIntArray;
  PartBValues: TDoubleArray;
  PartBKeyValue: Integer;
  PartBKeyTotal: Integer;
  AValue: Integer;
  APercent: double;
  Total: Double;
  TitleLine: string;
  RatingDate: string;
  EvidenceKeyArray: TIntArray;
  InterventionsKeyArray: TIntArray;
  OutcomesKeyArray: TIntArray;
//  EvidenceTypeChar: AnsiChar;
  EvidenceTypes: TStringList;
  CharIndex: AnsiChar;
  EvidenceTypeArray: TIntArray;
  MaxKey: Integer;
  EvidenceIndex: integer;
  AnArray: TIntArray;
  ErrorLabel: string;
  ProblemType: string;
  SheetName: string;
  SheetHasData: Boolean;
  MinimumPriorityScore: Integer;
  MaximumPriorityScore: Integer;
  PriorityScoreRange: Integer;
  MaxTier1: Integer;
  MaxTier2: Integer;
  MaxTier3: Integer;
  MinTier1: Integer;
  MinTier2: Integer;
  MinTier3: Integer;
  MinIndex: Integer;
//  MaxWrong: Integer;
  ProblemIDString: string;
  Splitter: TStringList;
  PriorityScoreFraction: double;
  EIO_Percent: Double;
  AveragePartB: double;
  TotalsLineBuilder: TStringBuilder;
  TotalsTitleLineBuilder: TStringBuilder;
  ALine: string;
  PartBPercent: double;
//  MaxEvidenceForParticipant: Integer;
  ProblemsIdentifiedArray: TIntArray;
//  MaxInterventionsForParticipant: integer;
//  MaxOutcomesForParticipant: Integer;
  AdjustedPriorityScore: Integer;
//  PossibleInterventions: TIntArray;
//  PossibleOutcomes: TIntArray;
begin
  SetLength(PartBKey, LastDataColPartB-FirstDataColPartB+1);
  SetLength(PartBValues, LastDataColPartB-FirstDataColPartB+1);
  SetLength(ProblemsIdentifiedArray, LastDataRowPartA-FirstDataRowPartA+1);
  InitializeScoringMatrix(ScoringMatrix);

  MaxTier1 := ScoringMatrix[0,0];
  MaxTier2 := ScoringMatrix[2,1];
  MaxTier3 := ScoringMatrix[6,2];
  MaximumPriorityScore := 2*MaxTier1 + 4*MaxTier2 + 4*MaxTier3;

  MinTier1 := ScoringMatrix[10,0];
  MinTier2 := ScoringMatrix[10,1];
  MinTier3 := ScoringMatrix[10,2];

  MinimumPriorityScore := 0;
//  MaxWrong := 0;
  // Compute penalties for Tier 1.
  for MinIndex := 9 downto 8 do
  begin
    MinTier1 := Min(ScoringMatrix[10,0], ScoringMatrix[MinIndex,0]);
//    if MinTier1 = ScoringMatrix[10,0] then
//    begin
//      Inc(MaxWrong);
//    end;
    MinimumPriorityScore := MinimumPriorityScore + MinTier1;
  end;
  // Compute penalties for Tier 2.
  for MinIndex := 7 downto 4 do
  begin
    MinTier2 := Min(ScoringMatrix[10,1], ScoringMatrix[MinIndex,1]);
//    if MinTier2 = ScoringMatrix[10,1] then
//    begin
//      Inc(MaxWrong);
//    end;
    MinimumPriorityScore := MinimumPriorityScore + MinTier2;
  end;
  // Compute penalties for Tier 3.
  for MinIndex := 3 downto 0 do
  begin
    MinTier3 := Min(ScoringMatrix[10,2], ScoringMatrix[MinIndex,2]);
//    if MinTier3 = ScoringMatrix[10,2] then
//    begin
//      Inc(MaxWrong);
//    end;
    MinimumPriorityScore := MinimumPriorityScore + MinTier3;
  end;
//  MinimumPriorityScore := MinimumPriorityScore -  MaxWrong;
  PriorityScoreRange := MaximumPriorityScore-MinimumPriorityScore;
  lblMax.Caption := 'Maximum Possible Priority Score = ' + IntToStr(MaximumPriorityScore);
  lblMin.Caption := 'Minimum Possible Priority Score = ' + IntToStr(MinimumPriorityScore);


  XLApp := CreateOleObject('Excel.Application');
  TotalsLineBuilder := TStringBuilder.Create;
  LineBuilder := TStringBuilder.Create;
  OutputFile := TStringList.Create;
  EvidenceTypes := TStringList.Create;
  TotalsTitleLineBuilder := TStringBuilder.Create;
  try
    for CharIndex := 'A' to 'Z' do
    begin
      EvidenceTypes.Add(Char(CharIndex))
    end;
    EvidenceTypes.CaseSensitive := True;
    XLApp.Visible := False;
    XLApp.Workbooks.Open(AXLSFile);
    try
      SheetCount := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets.Count;
      pbCompletion.Max := SheetCount;
      pbCompletion.Position := 0;
      Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];
      Sheet.Select;
      Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
      x := XLApp.ActiveCell.Row;
      y := XLApp.ActiveCell.Column;

      RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;

//      SetLength(PossibleEvidences, LastDataRowPartA+1);
      TotalEvidenceKey := 0;
      SetLength(EvidenceKeyArray, LastDataRowPartA-FirstDataRowPartA+1);
      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := RangeMatrix[RowIndex,EvidenceKeyCol];
        if AString = '' then
        begin
          AString := '0';
        end;
        Evidence := StrToInt(AString);
        EvidenceKeyArray[RowIndex-FirstDataRowPartA] := Evidence;
        TotalEvidenceKey := TotalEvidenceKey + Evidence;
      end;
      Sheet.Cells[TotalsPartARow, EvidenceKeyCol].Value := TotalEvidenceKey;

      TotalInterventionsKey := 0;
      SetLength(InterventionsKeyArray, LastDataRowPartA-FirstDataRowPartA+1);
      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := RangeMatrix[RowIndex,InterventionsKeyCol];
        if AString = '' then
        begin
          AString := '0';
        end;
        Interventions := StrToInt(AString);
        InterventionsKeyArray[RowIndex-FirstDataRowPartA] := Interventions;
        TotalInterventionsKey := TotalInterventionsKey + Interventions;
      end;
//      Sheet.Cells[TotalsPartARow, InterventionsKeyCol].Value := TotalInterventionsKey;
      // Lily decided to change the number of interventions to 21 for the purposes
      // of scoring but to leave 55 on the spreadsheet.
      TotalInterventionsKey := 21;

      TotalOutcomesKey := 0;
      SetLength(OutcomesKeyArray, LastDataRowPartA-FirstDataRowPartA+1);
      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := RangeMatrix[RowIndex,OutcomesKeyCol];
        if AString = '' then
        begin
          AString := '0';
        end;
        Outcomes := StrToInt(AString);
        OutcomesKeyArray[RowIndex-FirstDataRowPartA] := Outcomes;
        TotalOutcomesKey := TotalOutcomesKey + Outcomes;
      end;
      Sheet.Cells[TotalsPartARow, OutcomesKeyCol].Value := TotalOutcomesKey;

      Splitter := TStringList.Create;
      try
        PartBKeyTotal := 0;
        for ColIndex := FirstDataColPartB to LastDataColPartB do
        begin
          AString := RangeMatrix[PartBValueRowKey,ColIndex];
          Splitter.DelimitedText := AString;
          AString := Splitter[0];
          PartBKeyValue := StrToInt(AString);
          PartBKey[ColIndex-FirstDataColPartB] := PartBKeyValue;
          PartBKeyTotal := PartBKeyTotal + PartBKeyValue;
        end;
      finally
        Splitter.Free;
      end;

      // Start building line listing the variables.
      TotalsTitleLineBuilder.Append('StudyID, ');
//      TotalsTitleLineBuilder.Append('Rater, ');
//      TotalsTitleLineBuilder.Append('RatingDate, ');

      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := ProblemLabels[RowIndex-FirstDataRowPartA];
        LineBuilder.Append(AString);
        LineBuilder.Append('Present, ');
      end;

      TotalsTitleLineBuilder.Append('ProblemsIDdNum, ');
//      LineBuilder.Append('Problems_Identified_Percent, ');

      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := ProblemLabels[RowIndex-FirstDataRowPartA];
  //      AString := RangeMatrix[RowIndex,ProblemLabelCol];
        LineBuilder.Append(AString);
        LineBuilder.Append('Repeated, ');
      end;

      TotalsTitleLineBuilder.Append('RepeatedNum, ');

      TotalsTitleLineBuilder.Append('WrongNum, ');
//      LineBuilder.Append('NumWrong_Percent, ');

      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := ProblemLabels[RowIndex-FirstDataRowPartA];
        LineBuilder.Append(AString);
        LineBuilder.Append('Priority, ');
      end;

      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := ProblemLabels[RowIndex-FirstDataRowPartA];
        LineBuilder.Append(AString);
        LineBuilder.Append('PriorityPts, ');
      end;

      TotalsTitleLineBuilder.Append('PriorityScore, ');

//      LineBuilder.Append('Total_Priority_Score_Percent, ');

      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := ProblemLabels[RowIndex-FirstDataRowPartA];
        LineBuilder.Append(AString);
        LineBuilder.Append('Evidence, ');
        for EvidenceIndex := 0 to EvidenceKeyArray[RowIndex-FirstDataRowPartA] - 1 do
        begin
          LineBuilder.Append(AString);
          LineBuilder.Append('Evidence');
          LineBuilder.Append(EvidenceTypes[EvidenceIndex]);
          LineBuilder.Append(', ');
        end;
      end;
      TotalsTitleLineBuilder.Append('EvidenceNum, ');
      TotalsTitleLineBuilder.Append('EvidencePC, ');

      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := ProblemLabels[RowIndex-FirstDataRowPartA];
        LineBuilder.Append(AString);
        LineBuilder.Append('Interventions, ');
        for EvidenceIndex := 0 to InterventionsKeyArray[RowIndex-FirstDataRowPartA] - 1 do
        begin
          LineBuilder.Append(AString);
          LineBuilder.Append('Interventions');
          LineBuilder.Append(EvidenceTypes[EvidenceIndex]);
          LineBuilder.Append(', ');
        end;
      end;

      TotalsTitleLineBuilder.Append('InterventionsNum, ');
      TotalsTitleLineBuilder.Append('InterventionsPC, ');

      for RowIndex := FirstDataRowPartA to LastDataRowPartA do
      begin
        AString := ProblemLabels[RowIndex-FirstDataRowPartA];
        LineBuilder.Append(AString);
        LineBuilder.Append('Outcomes, ');
        for EvidenceIndex := 0 to OutcomesKeyArray[RowIndex-FirstDataRowPartA] - 1 do
        begin
          LineBuilder.Append(AString);
          LineBuilder.Append('Outcomes');
          LineBuilder.Append(EvidenceTypes[EvidenceIndex]);
          LineBuilder.Append(', ');
        end;
      end;
      TotalsTitleLineBuilder.Append('OutcomesNum, ');
      TotalsTitleLineBuilder.Append('OutcomesPC, ');

      for ColIndex := FirstDataColPartB to LastDataColPartB do
      begin
        AString := PartBLabels[ColIndex-FirstDataColPartB];
//        AString := StringReplace(AString, ' ', '', [rfReplaceAll, rfIgnoreCase]);
//        AString := StringReplace(AString, '_', '', [rfReplaceAll, rfIgnoreCase]);
        TotalsTitleLineBuilder.Append(AString);
        TotalsTitleLineBuilder.Append('Num, ');
        TotalsTitleLineBuilder.Append(AString);
        TotalsTitleLineBuilder.Append('PC, ');
        for EvidenceIndex := 0 to PartBKey[ColIndex-FirstDataColPartB] - 1 do
        begin
          LineBuilder.Append(AString);
          LineBuilder.Append('Identified');
          LineBuilder.Append(EvidenceTypes[EvidenceIndex]);
          LineBuilder.Append(', ');
        end;
      end;
//      for ColIndex := FirstDataColPartB to LastDataColPartB do
//      begin
//        AString := RangeMatrix[PartBProblemLabelRow,ColIndex];
//        AString := StringReplace(AString, ' ', '', [rfReplaceAll, rfIgnoreCase]);
//        AString := StringReplace(AString, '_', '', [rfReplaceAll, rfIgnoreCase]);
//      end;

      TotalsTitleLineBuilder.Append('PriorityPC, ');
      TotalsTitleLineBuilder.Append('EIOPC, ');

      TotalsTitleLineBuilder.Append('LMDPC, ');
      TotalsTitleLineBuilder.Append('CTScore, ');

//      LineBuilder.Append('Total');

      // TitleLine contains the variable names.
      TitleLine := TotalsTitleLineBuilder.ToString + Trim(LineBuilder.ToString);
      if TitleLine[Length(TitleLine)] = ',' then
      begin
        TitleLine := Copy(TitleLine, 1, Length(TitleLine)-1);
      end;
      pbCompletion.Position := 1;

//      Sheet := Sheet.Next;
      // The first sheet is the template so start on sheet 2.
      for SheetIndex := 2 to SheetCount do
      begin
        TotalsLineBuilder.Clear;
        LineBuilder.Clear;
        SheetHasData := False;
//        Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[SheetIndex];
        Sheet := XLApp.WorkSheets[SheetIndex];
        SheetName := Sheet.Name;
        try
          Sheet.Select;
        except on EOleException do
          begin
            memoErrors.Lines.Add(Format('Sheet "%s" is hidden and will be skipped.', [SheetName]));
            Continue;
          end;
        end;
        Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;

        Sheet.Cells[TotalsPartARow, EvidenceKeyCol].Value := TotalEvidenceKey;
        // Lily decided to leave the value of 55 on the spreadsheet but to use
        // 23 in the calculation of the interventions percent.
//        Sheet.Cells[TotalsPartARow, InterventionsKeyCol].Value := TotalInterventionsKey;
        Sheet.Cells[TotalsPartARow, OutcomesKeyCol].Value := TotalOutcomesKey;

        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;

        ParticipantID := RangeMatrix[IDRow,IDCol];
        if ParticipantID = '' then
        begin
          memoErrors.Lines.Add('No Participant ID on sheet ' + SheetName);
        end;
        TotalsLineBuilder.Append(ParticipantID);
        TotalsLineBuilder.Append(', ');

//        RaterID := RangeMatrix[IDRow,RaterIDCol];
//        TotalsLineBuilder.Append(RaterID);
//        TotalsLineBuilder.Append(', ');
//
//        RatingDate := RangeMatrix[IDRow,DateCol];
//        TotalsLineBuilder.Append(RatingDate);
//        TotalsLineBuilder.Append(', ');
//
        TotalProblems := 0;
        for RowIndex := FirstDataRowPartA to LastDataRowPartA do
        begin
          AString := RangeMatrix[RowIndex,ProblemPresentCol];
          if AString = '' then
          begin
            AString := '0';
          end
          else
          begin
            SheetHasData := True;
          end;
          LineBuilder.Append(AString);
          LineBuilder.Append(', ');
          Present := StrToInt(AString);
          if not (Present in [0,1]) then
          begin
            memoErrors.Lines.Add(Format(
              'Incorrect coding for "problem present" for %s on sheet %s',
                [ParticipantID, SheetName]))
          end;
          ProblemsIdentifiedArray[RowIndex-FirstDataRowPartA] := Present;
          TotalProblems := TotalProblems + Present;
        end;
//        if TotalProblems > 0 then
//        begin
//          SheetHasData := True;
//        end;
    //  Total Problems Identified
        TotalsLineBuilder.Append(TotalProblems);
        TotalsLineBuilder.Append(', ');
        Sheet.Cells[TotalsPartARow, ProblemPresentCol].Value := TotalProblems;

    //  Percent Problems Identified);
//        PercentProblems := TotalProblems/10;
//        LineBuilder.Append(PercentProblems);
//        LineBuilder.Append(', ');
//        Sheet.Cells[PercentPartARow, ProblemPresentCol].NumberFormat := '0%';
//        Sheet.Cells[PercentPartARow, ProblemPresentCol].Value := PercentProblems;

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
        TotalsLineBuilder.Append(TotalRepeated);
        TotalsLineBuilder.Append(', ');
        Sheet.Cells[TotalsPartARow, ProblemRepeatedCol].Value := TotalRepeated;

  //    Unkeyed Score
        AString := RangeMatrix[TotalsPartARow,WrongProblemCol];
        Sheet.Cells[RepeatWrongProbRow, RepeatWrongProbCol].Value := AString;
        AString := UpperCase(AString);

        if (AString = '') then
        begin
          AString := '0';
          if (ParticipantID <> '') then
          begin
            memoErrors.Lines.Add(Format('In participant %s on sheet %s, the number of wrong problems was not specified',
              [ParticipantID, SheetName]));
          end;
        end;
        WrongProbs := StrToInt(AString);
        TotalsLineBuilder.Append(WrongProbs);
        TotalsLineBuilder.Append(', ');
//        Sheet.Cells[TotalsPartARow, WrongProblemCol].Value := WrongProbs;

  //    Unkeyed Score Percent
//        PercentWrongProbs := (10-WrongProbs)/10;
//        LineBuilder.Append(PercentWrongProbs);
//        LineBuilder.Append(', ');
//        Sheet.Cells[PercentPartARow, WrongProblemCol].NumberFormat := '0%';
//        Sheet.Cells[PercentPartARow, WrongProblemCol].Value := PercentWrongProbs;

        for RowIndex := FirstDataRowPartA to LastDataRowPartA do
        begin
          ProblemIDString := RangeMatrix[RowIndex,ProblemPresentCol];
          AString := RangeMatrix[RowIndex,AssignedPriorityCol];
          if AString = '' then
          begin
            Priority := 0;
            if ProblemIDString <> '' then
            begin
              memoErrors.Lines.Add(Format('In participant %s on sheet %s, a priority was specified incorrectly',
                [ParticipantID, SheetName]));
            end;
          end
          else
          begin
            Priority := StrToInt(AString);
            if ProblemIDString = '' then
            begin
              memoErrors.Lines.Add(Format('In participant %s on sheet %s, a priority was specified incorrectly',
                [ParticipantID, SheetName]));
            end;
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
          Tier := -1;
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

//        PriorityScoreTotal := PriorityScoreTotal;// - MinimumPriorityScore;

        Sheet.Cells[TotalsPartARow, PriorityScoreCol].Value := PriorityScoreTotal;

        AdjustedPriorityScore := PriorityScoreTotal - MinimumPriorityScore;
        Sheet.Cells[AdjustedPriorityPointsRow, PriorityScoreCol].Value := AdjustedPriorityScore;

        PriorityScoreTotal := AdjustedPriorityScore - WrongProbs;
        TotalsLineBuilder.Append(PriorityScoreTotal);
        TotalsLineBuilder.Append(', ');
        Sheet.Cells[PrioritySubScoreRow, PriorityScoreCol].Value := PriorityScoreTotal;


        PriorityScoreFraction := PriorityScoreTotal/PriorityScoreRange;
        Sheet.Cells[PriorityPercentRow, PriorityScoreCol].Value := PriorityScoreFraction;
//        Sheet.Cells[TotalPriorityScoreRow, PriorityScoreCol].NumberFormat := '0%';

//        PriorityScorePercent := (PriorityScoreTotal-MinimumPriorityScore)
//          /PriorityScoreRange;
//        LineBuilder.Append(PriorityScorePercent);
//        LineBuilder.Append(', ');
//        Sheet.Cells[PercentPartARow, PriorityScoreCol].NumberFormat := '0%';
//        Sheet.Cells[PercentPartARow, PriorityScoreCol].Value := PriorityScorePercent;

        TotalEvidence := 0;
//        MaxEvidenceForParticipant := 0;
        for RowIndex := FirstDataRowPartA to LastDataRowPartA do
        begin
          MaxKey := EvidenceKeyArray[RowIndex-FirstDataRowPartA];
//          MaxEvidenceForParticipant := MaxEvidenceForParticipant
//            + ProblemsIdentifiedArray[RowIndex - FirstDataRowPartA]*MaxKey;
          InitializeEvidenceArray(EvidenceTypeArray, MaxKey);
          AString := RangeMatrix[RowIndex,EvidenceCol];
          AString := Trim(AString);
          ProblemIDString := RangeMatrix[RowIndex,ProblemPresentCol];
          if ProblemIDString = '' then
          begin
            if AString <> '' then
            begin
              memoErrors.Lines.Add(Format('In participant %s on sheet %s, evidence was specified incorrectly',
                [ParticipantID, SheetName]));
            end;
          end
          else
          begin
            if AString = '' then
            begin
              memoErrors.Lines.Add(Format('In participant %s on sheet %s, evidence was specified incorrectly',
                [ParticipantID, SheetName]));
            end;
          end;
          ErrorLabel := 'evidence';
          ProblemType := RangeMatrix[RowIndex, ProblemLabelCol];
          ReadEvidenceTypes(EvidenceTypeArray, ParticipantID, AString,
            EvidenceTypes, ErrorLabel, ProblemType, SheetName);
          CountEvidence(LineBuilder, EvidenceTypeArray, Evidence, ErrorLabel,
            ParticipantID, ProblemType, SheetName);
//          Sheet.Cells[RowIndex, EvidencePercentCol].NumberFormat := '0%';
//          Sheet.Cells[RowIndex, EvidencePercentCol].Value := Evidence/MaxKey;
          Sheet.Cells[RowIndex, EvidencePercentCol].Value := Evidence/MaxKey;
          StoreEvidenceTypeValues(LineBuilder, EvidenceTypeArray);
          TotalEvidence := TotalEvidence + Evidence;
        end;

  //    Total Evidence
        TotalsLineBuilder.Append(TotalEvidence);
        TotalsLineBuilder.Append(', ');
        Sheet.Cells[TotalsPartARow, EvidenceCol].Value := TotalEvidence;

  //    Evidence Percent
//        if MaxEvidenceForParticipant <> 0 then
//        begin
          TotalEvidencePercent := TotalEvidence/TotalEvidenceKey;
//        end
//        else
//        begin
//          TotalEvidencePercent := 0;
//        end;
        TotalsLineBuilder.Append(TotalEvidencePercent);
        TotalsLineBuilder.Append(', ');
//        Sheet.Cells[PercentPartARow, EvidenceCol].NumberFormat := '0%';
        Sheet.Cells[TotalsPartARow, EvidencePercentCol].Value := TotalEvidencePercent;

        TotalInterventions := 0;
//        MaxInterventionsForParticipant := 0;
        for RowIndex := FirstDataRowPartA to LastDataRowPartA do
        begin
  //        AString := RangeMatrix[RowIndex,InterventionsCol];
  //        if AString = '' then
  //        begin
  //          AString := '0';
  //        end;
  //        LineBuilder.Append(AString);
  //        LineBuilder.Append(', ');
  //        Intervetions := StrToInt(AString);
  //        TotalInterventions := TotalInterventions + Intervetions;

          MaxKey := InterventionsKeyArray[RowIndex-FirstDataRowPartA];
//          MaxInterventionsForParticipant := MaxInterventionsForParticipant
//            + ProblemsIdentifiedArray[RowIndex - FirstDataRowPartA]*MaxKey;
          InitializeEvidenceArray(EvidenceTypeArray, MaxKey);
          AString := RangeMatrix[RowIndex,InterventionsCol];
          AString := Trim(AString);
          ProblemIDString := RangeMatrix[RowIndex,ProblemPresentCol];
          if ProblemIDString = '' then
          begin
            if AString <> '' then
            begin
              memoErrors.Lines.Add(Format('In participant %s on sheet %s, interventions was specified incorrectly',
                [ParticipantID, SheetName]));
            end;
          end
          else
          begin
            if AString = '' then
            begin
              memoErrors.Lines.Add(Format('In participant %s on sheet %s, interventions was specified incorrectly',
                [ParticipantID, SheetName]));
            end;
          end;
          ErrorLabel := 'interventions';
          ProblemType := RangeMatrix[RowIndex, ProblemLabelCol];
          ReadEvidenceTypes(EvidenceTypeArray, ParticipantID, AString,
            EvidenceTypes, ErrorLabel, ProblemType, SheetName);
          CountEvidence(LineBuilder, EvidenceTypeArray, Intervetions, ErrorLabel,
            ParticipantID, ProblemType, SheetName);
//          Sheet.Cells[RowIndex, InterventionsPercentCol].NumberFormat := '0%';
//          Sheet.Cells[RowIndex, InterventionsPercentCol].Value := Intervetions/MaxKey;
          Sheet.Cells[RowIndex, InterventionsPercentCol].Value := Intervetions/MaxKey;
          StoreEvidenceTypeValues(LineBuilder, EvidenceTypeArray);
          TotalInterventions := TotalInterventions + Intervetions;

        end;

        TotalsLineBuilder.Append(TotalInterventions);
        TotalsLineBuilder.Append(', ');
        Sheet.Cells[TotalsPartARow, InterventionsCol].Value := TotalInterventions;

        // InterventionsPercent
//        if MaxInterventionsForParticipant <> 0 then
//        begin
          TotalInterventionsPercent := TotalInterventions/TotalInterventionsKey;
//        end
//        else
//        begin
//          TotalInterventionsPercent := 0;
//        end;
        TotalsLineBuilder.Append(TotalInterventionsPercent);
        TotalsLineBuilder.Append(', ');
//        Sheet.Cells[PercentPartARow, InterventionsCol].NumberFormat := '0%';
        Sheet.Cells[TotalsPartARow, InterventionsPercentCol].Value := TotalInterventionsPercent;

        TotalOutcomes := 0;
//        MaxOutcomesForParticipant := 0;
        for RowIndex := FirstDataRowPartA to LastDataRowPartA do
        begin
  //        AString := RangeMatrix[RowIndex,OutcomesCol];
  //        if AString = '' then
  //        begin
  //          AString := '0';
  //        end;
  //        LineBuilder.Append(AString);
  //        LineBuilder.Append(', ');
  //        Outcomes := StrToInt(AString);
  //        TotalOutcomes := TotalOutcomes + Outcomes;

          MaxKey := OutcomesKeyArray[RowIndex-FirstDataRowPartA];
//          MaxOutcomesForParticipant := MaxOutcomesForParticipant
//            + ProblemsIdentifiedArray[RowIndex - FirstDataRowPartA]*MaxKey;
          InitializeEvidenceArray(EvidenceTypeArray, MaxKey);
          AString := RangeMatrix[RowIndex,OutcomesCol];
          AString := Trim(AString);
          ProblemIDString := RangeMatrix[RowIndex,ProblemPresentCol];
          if ProblemIDString = '' then
          begin
            if AString <> '' then
            begin
              memoErrors.Lines.Add(Format('In participant %s on sheet %s, outcomes was specified incorrectly',
                [ParticipantID, SheetName]));
            end;
          end
          else
          begin
            if AString = '' then
            begin
              memoErrors.Lines.Add(Format('In participant %s on sheet %s, outcomes was specified incorrectly',
                [ParticipantID, SheetName]));
            end;
          end;
          ErrorLabel := 'outcomes';
          ProblemType := RangeMatrix[RowIndex, ProblemLabelCol];
          ReadEvidenceTypes(EvidenceTypeArray, ParticipantID, AString,
            EvidenceTypes, ErrorLabel, ProblemType, SheetName);
          CountEvidence(LineBuilder, EvidenceTypeArray, Outcomes, ErrorLabel,
            ParticipantID, ProblemType, SheetName);
//          Sheet.Cells[RowIndex, OutcomesPercentCol].NumberFormat := '0%';
//          Sheet.Cells[RowIndex, OutcomesPercentCol].Value := Outcomes/MaxKey;
          Sheet.Cells[RowIndex, OutcomesPercentCol].Value := Outcomes/MaxKey;
          StoreEvidenceTypeValues(LineBuilder, EvidenceTypeArray);
          TotalOutcomes := TotalOutcomes + Outcomes;
        end;
        Sheet.Cells[TotalsPartARow, OutcomesCol].Value := TotalOutcomes;

        TotalsLineBuilder.Append(TotalOutcomes);
        TotalsLineBuilder.Append(', ');

        // OutcomesPercent
//        if MaxOutcomesForParticipant <> 0 then
//        begin
          TotalOutcomesPercent := TotalOutcomes/TotalOutcomesKey;
//        end
//        else
//        begin
//          TotalOutcomesPercent := 0;
//        end;
        TotalsLineBuilder.Append(TotalOutcomesPercent);
        TotalsLineBuilder.Append(', ');
//        Sheet.Cells[PercentPartARow, OutcomesCol].NumberFormat := '0%';
        Sheet.Cells[TotalsPartARow, OutcomesPercentCol].Value := TotalOutcomesPercent;

        EIO_Percent := (TotalEvidencePercent + TotalInterventionsPercent
          + TotalOutcomesPercent)/3;

        Sheet.Cells[PriorityPercentRow, InterventionsPercentCol].Value := EIO_Percent;


        for ColIndex := FirstDataColPartB to LastDataColPartB do
        begin
  //        AString := RangeMatrix[PartBValueRow,ColIndex];
  //        if AString = '' then
  //        begin
  //          AString := '0';
  //        end;
  //        LineBuilder.Append(AString);
  //        LineBuilder.Append(', ');
  //        AValue := StrToInt(AString);
  //        PartBValues[ColIndex-FirstDataColPartB] := AValue;

          MaxKey := PartBKey[ColIndex-FirstDataColPartB];
          InitializeEvidenceArray(EvidenceTypeArray, MaxKey);
          AString := RangeMatrix[PartBValueRow,ColIndex];
          AString := Trim(AString);
          ErrorLabel := 'Part B';
          ProblemType := RangeMatrix[PartBLbaelRow, ColIndex];
          if (AString = '') and (ParticipantID <> '') then
          begin
            memoErrors.Lines.Add(Format('incorrect coding for %s for Participant %s on sheet %s', [ProblemType, ParticipantID, sheetName]));
          end;
          ReadEvidenceTypes(EvidenceTypeArray, ParticipantID, AString,
            EvidenceTypes, ErrorLabel, ProblemType, SheetName);
          CountEvidence(TotalsLineBuilder, EvidenceTypeArray, AValue, ErrorLabel,
            ParticipantID, ProblemType, SheetName);
//          Sheet.Cells[PartBValuePercentRow, ColIndex].NumberFormat := '0%';
//          Sheet.Cells[PartBValuePercentRow, ColIndex].Value := AValue/MaxKey;
          Sheet.Cells[PartBValuePercentRow, ColIndex].Value := AValue/MaxKey;
          StoreEvidenceTypeValues(LineBuilder, EvidenceTypeArray);
          PartBPercent := AValue/MaxKey;
          PartBValues[ColIndex-FirstDataColPartB] := PartBPercent;
          TotalsLineBuilder.Append(PartBPercent);
          TotalsLineBuilder.Append(', ');
        end;

        AveragePartB := 0.;
        for ColIndex := 0 to Length(PartBValues) - 1 do
        begin
          AveragePartB := AveragePartB + PartBValues[ColIndex];
//          TotalsLineBuilder.Append(PartBValues[ColIndex]);
//          TotalsLineBuilder.Append(', ');
        end;
        AveragePartB := AveragePartB/Length(PartBValues);
        Sheet.Cells[PriorityPercentRow, Succ(FirstDataColPartB)].Value := AveragePartB;

        TotalsLineBuilder.Append(PriorityScoreFraction);
        TotalsLineBuilder.Append(', ');

        TotalsLineBuilder.Append(EIO_Percent);
        TotalsLineBuilder.Append(', ');

        TotalsLineBuilder.Append(AveragePartB);
        TotalsLineBuilder.Append(', ');

        Total := 0.4*PriorityScoreFraction + 0.45*EIO_Percent + 0.15*AveragePartB;
        TotalsLineBuilder.Append(Total);
        TotalsLineBuilder.Append(', ');
        Sheet.Cells[TotalRow, TotalCol].Value := Total;

//        Total := PercentProblems + PercentWrongProbs + PriorityScorePercent
//          + TotalEvidencePercent + TotalInterventionsPercent + TotalOutcomesPercent;
//        for ColIndex := FirstDataColPartB to LastDataColPartB do
//        begin
//          APercent := PartBValues[ColIndex-FirstDataColPartB]/
//            PartBKey[ColIndex-FirstDataColPartB];
//          LineBuilder.Append(APercent);
//          LineBuilder.Append(', ');
//          Sheet.Cells[PartBValuePercentRow, ColIndex].NumberFormat := '0%';
//          Sheet.Cells[PartBValuePercentRow, ColIndex].Value := APercent;
//
//          Total := Total + APercent;
//        end;
//        Total := Total / 9;
//        LineBuilder.Append(Total);
//        Sheet.Cells[TotalRow, TotalCol].NumberFormat := '0%';
//        Sheet.Cells[TotalRow, TotalCol].Value := Total;

        Sheet.Range['A1','A1'].Select;

        RangeMatrix := Unassigned;
        ALine := Trim(TotalsLineBuilder.ToString + LineBuilder.ToString);
        if ALine[Length(ALine)] = ',' then
        begin
          ALine := Copy(ALine, 1, Length(ALine)-1);
        end;
        OutputFile.Add(ALine);
        pbCompletion.Position := SheetIndex;

        if SheetHasData and (ParticipantID = '') then
        begin
          memoErrors.Lines.Add('ID not specified on sheet ' + SheetName)
        end;

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
      XLApp.DisplayAlerts := False;
      XLApp.Workbooks[ExtractFileName(AXLSFile)].Close
    end;
  finally
    TotalsLineBuilder.Free;
    TotalsTitleLineBuilder.Free;
    LineBuilder.Free;
    OutputFile.Free;
    EvidenceTypes.Free;
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
      ScoringMatrix[ColIndex - 1, RowIndex - 1] :=
        rdgScoringMatrix.IntegerValue[ColIndex, RowIndex];
    end;
  end;
end;

procedure TfrmMain.StoreEvidenceTypeValues(LineBuilder: TStringBuilder;
  AnArray: TIntArray);
var
  CharTypeIndex: Integer;
begin
  for CharTypeIndex := 0 to Length(AnArray) - 1 do
  begin
    if (AnArray[CharTypeIndex] > 0) then
    begin
      LineBuilder.Append(1);
    end
    else
    begin
      LineBuilder.Append(0);
    end;
    LineBuilder.Append(', ');
  end;
end;

procedure TfrmMain.CountEvidence(LineBuilder: TStringBuilder; AnArray: TIntArray;
  var Evidence: Integer; const ErrorLabel, ParticipantID, ProblemType, SheetName: string);
var
  CharTypeIndex: Integer;
begin
  Evidence := 0;
  for CharTypeIndex := 0 to Length(AnArray) - 1 do
  begin
    if (AnArray[CharTypeIndex] > 0) then
    begin
      Inc(Evidence);
      if AnArray[CharTypeIndex] > 1 then
      begin
        memoErrors.Lines.Add(Format('duplicate coding for "%s" for %s in %s on Sheet %s.',
          [ErrorLabel, ProblemType, ParticipantID, SheetName]));
      end;
    end;
  end;
  LineBuilder.Append(Evidence);
  LineBuilder.Append(', ');
end;

procedure TfrmMain.InitializeEvidenceArray(var AnArray: TIntArray; MaxKey: integer);
var
  CharTypeIndex: Integer;
begin
  SetLength(AnArray, MaxKey);
  for CharTypeIndex := 0 to Length(AnArray) - 1 do
  begin
    AnArray[CharTypeIndex] := 0;
  end;
end;

procedure TfrmMain.ReadEvidenceTypes(AnArray: TIntArray; ParticipantID: string;
   AString: string; EvidenceTypes: TStringList;
   const ErrorLabel, ProblemType, SheetName: string);
var
  AChar: Char;
  CharTypeIndex: Integer;
  CharOrdinal: Integer;
  ErrorFound: Boolean;
begin
  // X means no evidence
  if AString = 'X' then
  begin
    Exit;
  end;

  ErrorFound := False;
  for CharTypeIndex := 1 to Length(AString) do
  begin
    AChar := AString[CharTypeIndex];
    CharOrdinal := EvidenceTypes.IndexOf(AChar);
    if (CharOrdinal < 0) or (CharOrdinal >= Length(AnArray)) then
    begin
      ErrorFound := True;
    end
    else
    begin
      Inc(AnArray[CharOrdinal]);
    end;
  end;
  if ErrorFound then
  begin
    memoErrors.Lines.Add(Format('Incorrect coding for "%s" for %s for participant "%s" on sheet %s',
      [ErrorLabel, ProblemType, ParticipantID, SheetName]));
  end;
end;

end.
