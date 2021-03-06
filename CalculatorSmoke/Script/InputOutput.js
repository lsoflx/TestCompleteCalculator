﻿//USEUNIT ExcelFunc

function GetNumber(string)
//Extracts number from value
{
 return string.replace(/[^\d.\/*+-]/g, '');
}

function SysAndMapRefresh()
//Refreshes Sys.Processes and NameMapping for getting displayed result
{
  //Refreshes Sys.Processes
  Sys.Process('Microsoft.WindowsCalculator').Refresh();
  //Refreshes NameMapping
  Aliases.Microsoft_WindowsCalculator.result.RefreshMappingInfo();
}

function GetRow()
//Connects rows to ODT for independent count
{
 //Returns current row
 return ODT.Data.CalcGroup.Calc.Row;
}

function CountRow()
//Counts rows
{
  ODT.Data.CalcGroup.Calc.Row = String(parseInt(ODT.Data.CalcGroup.Calc.Row) + 1);
}

function GetInputValue()
//Returns input expression
{
  //Declares varialbe of input data
  typeIn = GetNumber(String(Excel.Cells.Item(GetRow() , 4)));
  //Input message goes to log
  Log.Message(typeIn);
  return typeIn;
}

function GetExpectedValue()
//Returns expected result
{
  //Returns expected result
  return GetNumber(String(Excel.Cells.Item(GetRow(), 5)));
}

function InputExpression(expression)
//Inputs string 'expression' with keyboard into calculator
{
  //Grabs calculator window 
  wCalc = Aliases.Microsoft_WindowsCalculator.Calculator;
  //Goes through every char of 'expression'
  for(i = 0; i < expression.length; i++)
  {
    //Inputs 'expression' with keyboard
    wCalc.Keys(expression.substr(i, 1));
  }
  //Inputs '=' for getting result
  wCalc.Keys('=');
}

function GetDislpayResult()
//Returns number on display
{
  //Refreshes Sys.Processes and NameMapping for getting displayed result
  SysAndMapRefresh();
  //Grabs and returns the result from display
  return GetNumber(String(Aliases.Microsoft_WindowsCalculator.result.Name));
}

function CompareActualExpectedResult(expectedOutput, displayed)
//Check if expected result is equal to actual
{
  aqObject.CompareProperty(displayed, cmpEqual, expectedOutput);
  //Counts rows
  CountRow();
}

//Exports functions to MainTest
module.exports.GetInputValue = GetInputValue;
module.exports.GetExpectedValue = GetExpectedValue;
module.exports.InputExpression = InputExpression;
module.exports.GetDislpayResult = GetDislpayResult;
module.exports.CompareActualExpectedResult = CompareActualExpectedResult;