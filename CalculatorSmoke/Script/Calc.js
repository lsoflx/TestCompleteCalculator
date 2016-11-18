﻿function _CalcStartODT()
{
  TestedApps.calc.Run();
  var wCalc = Sys.Process('Microsoft.WindowsCalculator').Window('Windows.UI.Core.CoreWindow', 'Calculator', 1);
  var display = String(Aliases.Microsoft_WindowsCalculator.result.Name);
  ODT.Data.CalcGroup.Calc.Wnd = wCalc;
  ODT.Data.CalcGroup.Calc.Result = display.replace(/[^\d.-]/g, '');
}

function _CalcStopODT()
{
  Sys.Process('Microsoft.WindowsCalculator').Close();
}

function _CalcCalculateODT(expression)
{
  var i;
  var wCalc = Sys.Process('Microsoft.WindowsCalculator').Window('Windows.UI.Core.CoreWindow', 'Calculator', 1);
  for(i = 0; i < expression.length; i++)
  {
    wCalc.Keys(expression.substr(i, 1));
  }
  wCalc.Keys('=');
  Sys.Process('Microsoft.WindowsCalculator').Refresh()
  Aliases.Microsoft_WindowsCalculator.result.RefreshMappingInfo()
  var display = String(Aliases.Microsoft_WindowsCalculator.result.Name);
  var result = String(display).replace(/[^\d.-]/g, '');
  ODT.Data.CalcGroup.Calc.Result = String(display).replace(/[^\d.-]/g, '')
  return result;
}

function ReadExcel()
{
  var calc = ODT.Data.CalcGroup.Calc;
  calc.Start();
  Excel = Sys.OleObject("Excel.Application");
  Excel.Workbooks.Open("C:\\CalcTestCasesSmoke.xlsx");

  Row = ODT.Data.CalcGroup.Calc.Row;
  type_in = (VarToString(Excel.Cells.Item(Row, 4)).replace(/[^\d.\/*+-]/g, ''));
  expected_result = (VarToString(Excel.Cells.Item(Row, 5)).replace(/[^\d.-]/g, ''));
  Log.Message(type_in);
  calc.Calculate(type_in);
  aqObject.CompareProperty(ODT.Data.CalcGroup.Calc.Result,cmpEqual, expected_result);
  calc.Close()
  
  ODT.Data.CalcGroup.Calc.Row = String(parseInt(ODT.Data.CalcGroup.Calc.Row) + 1);
  Excel.Quit();
}



function SearchRow(num, row)
{
  let Excel = Sys.OleObject("Excel.Application");
  Excel.Workbooks.Open("C:\\CalcTestCasesSmoke.xlsx");

  let RowCount = Excel.ActiveSheet.UsedRange.Rows.Count;

  for (let i = 1; i <= RowCount; i++)
  {
    let s = "";
    s += (VarToString(Excel.Cells.Item(i, row)) + "\r\n");
    if (aqString.StrMatches(num, s)){ODT.Data.CalcGroup.Calc.Row = i; break}; 
  }
  Excel.Quit();
}

function SearchPlus()
{
  SearchRow('Plus', 3);
}

function SearchMinus()
{
  SearchRow('Minus', 3);
}

function SearchMultiple()
{
  SearchRow('Multiple', 3);
}

function SearchDivide()
{
  SearchRow('Divide', 3);
}

function SearchCritical()
{
  SearchRow('Critical', 2);
}
function SearchMajor()
{
  SearchRow('Major', 2);
}