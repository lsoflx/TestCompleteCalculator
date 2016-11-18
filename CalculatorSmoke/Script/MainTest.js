//USEUNIT ExcelFunc
//USEUNIT InputOutput
//USEUNIT StartStopApp

function MainTest()
//Calculate and check result
{
  //Opens Calculator
  Start('Microsoft.WindowsCalculator', 0);
  //Connects Excel
  ConnectExcel(GetFilePath())
  //Inputs values
  InputExpression(GetInputValue());
  //Compares Actual and expected result
  CompareActualExpectedResult(GetExpectedValue(), GetDislpayResult());
  //Disconnects Excel
  Excel.Quit();
  //Closes calculator
  Stop('Microsoft.WindowsCalculator');
}