﻿
function CheckProcessExists(process)
//Check existance of process and closes it
{
  //Checks if process exists already
  if (Sys.WaitProcess(process).Exists)
  {
    //Closes process if it's opened
    Sys.Process(process).Close()
  }
  //Continues execution
  else;
}

function Start(process, appindex)
//Opens tested app
{ 
  //Checks if app opened and closes it
  CheckProcessExists(process);
  //Runs testedapp
  TestedApps.Items(appindex).Run();
}

function Stop(process)
//Closes testded app
{ 
  //Checks if app opened and closes it
  CheckProcessExists(process);
}

//Exports functions to MainTest 
module.exports.CheckProcessExists = CheckProcessExists;
module.exports.Start = Start;
module.exports.Stop = Stop;