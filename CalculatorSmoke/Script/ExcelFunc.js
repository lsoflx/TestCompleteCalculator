﻿//USEUNIT StartStopApp

function GetFilePath()
//Returns path of excel file
{
 return Project.Path + 'CalcTestCasesSmoke.xlsx';
}

function ConnectExcel(path)
//Connects excel
{
  //Checks if process running and closes it
  CheckProcessExists('Excel.Application');
  //Declares variaible for Excel app
  Excel = Sys.OleObject("Excel.Application");
  //Connects excel file
  Excel.Workbooks.Open(path);
}

function SearchRow(name, row)
//Search rows with given text
{
  //Connects excel
  ConnectExcel(GetFilePath())
  //Defines max rows amount
  RowCount = Excel.ActiveSheet.UsedRange.Rows.Count;
  //Parse through file
  for (let i = 1; i <= RowCount; i++)
  {
    //Variable contains cell text
    s = "";
    s += (VarToString(Excel.Cells.Item(i, row)) + "\r\n");
    //Finds match in cell text and 'name' text
    if (aqString.StrMatches(name, s)){ODT.Data.CalcGroup.Calc.Row = i; break}; 
  }
  //Disconnects excel
  Excel.Quit();
}

//Exports functions to MainTest and Navigation
module.exports.GetFilePath = GetFilePath;
module.exports.ConnectExcel = ConnectExcel;
module.exports.SearchRow = SearchRow;