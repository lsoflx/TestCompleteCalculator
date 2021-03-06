﻿//USEUNIT ExcelFunc
//USEUNIT InputOutput
//USEUNIT StartStopApp

function ManualTesting()
{
  Start('Microsoft.WindowsCalculator', 0)
  InputExpression('4+3')
  aqObject.CompareProperty(DislpayResult(),cmpEqual, '7')
  Stop('Microsoft.WindowsCalculator')
}

function GetRelativePathExample()
{
  // Specifies the current folder
  var sFolderPath = "C:\\Program Files (x86)\\SmartBear\\TestComplete 12\\Bin"; 
  // Specifies the fully qualified file name
  var sFilePath = "C:\\Users\\rsoroka\\Documents\\TestComplete 12 Projects\\CalculatorTest\\CalcTestCasesSmoke.xlsx";
  
  // Obtains the file's relative path
  var sRelPath = aqFileSystem.GetRelativePath(sFolderPath, sFilePath);
  
  // Posts the relative path to the test log
  Log.Message(sRelPath);
}