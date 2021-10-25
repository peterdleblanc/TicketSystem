/*
	~--------------------------------------------------~
	~NO EDITING OF THIS DOCUMENT IS ALLOWED.
	~PLEASE RESPECT YOUR SYSTEM ADMINISTRATOR AND FOLLOW
	THE LAYED-OUT RULES WHEN USING THIS PROGRAM.
	~--------------------------------------------------~
	
	~ (c)2007 ~
	Start Date; 10.12.07
	Change Date: 11.05.07
	**********************************************
	Database Backend Programming by: Peter LeBlanc
	Layout Programmed by: Craig Sheppard
*/

//.................................Connection String Variables..................................//
  var iStr;
  var iStr2;
  var iStr_login;
  var iStr_usermanagement;
  
  var adOpenKeyset = 1;
  var adLockOptimistic = 2;
  var adCmdText = 1;
  var adUseClient = 3;
//....................................ActiveX Variables.....................................//
  var cObj;
  var cObj2;
  var cObj_login;
  var cObj_usermanagement;
  var cObjCSV
  
  var recSet;
  var recSet2;
  var recSet3;
  var recSet4;
  var recSet_login;
  var recSet_usermanagement;
  var recSetCSV
  
  var grStr;
  var grStr2;
  var grStr3;
  var grStr4;
  var grStr_login;
  var grStr_usermanagement;
  
  var TestVariable;

//.................................Connection String Function.................................//

function SetupActiveXObjectCSV()
{
   cObjCSV = new ActiveXObject("ADODB.Connection");
   recSetCSV = new ActiveXObject("ADODB.Recordset");

   recSetCSV.CursorLocation = adUseClient;
   cObjCSV.CursorLocation = adUseClient;

}


function SetupConnectionStringCSV()
{
    	
	iStr='Provider=MSDASQL; Driver={Microsoft Text Driver (*.txt; *.csv)}; DBQ=T:\\1.5 Help Desk\\';
   
}	

function OpenDatabaseCSV()
{
   SetupConnectionStringCSV()
   SetupActiveXObjectCSV()

   cObjCSV.Open(iStr)
      
}

function OpenTableCSV()
{
   recSetCSV.Open(grStr,cObjCSV,adOpenKeyset,adLockOptimistic,adCmdText)

}

function SetupConnectionString()
{
   iStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\\1.5 Help Desk\\pub\\databases\\DefaultTickets_revised.mdb;Persist Security Info=False";
}

function SetupConnectionString2()
{
   iStr2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\\1.5 Help Desk\\pub\\databases\\Stats.mdb;Persist Security Info=False";
}

function SetupConnectionString_login()
{
   iStr_login = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\\All In One\\Reports Viewer\\pub\\databases\\UserInfo.mdb;Persist Security Info=False";
}

//...................................ActiveX Setup Function....................................//
function SetupActiveXObject()
{
   cObj = new ActiveXObject("ADODB.Connection");
   recSet = new ActiveXObject("ADODB.Recordset");
   recSet2 = new ActiveXObject("ADODB.Recordset");
   recSet4 = new ActiveXObject("ADODB.Recordset");

   recSet.CursorLocation = adUseClient;
   recSet2.CursorLocation = adUseClient;
   recSet4.CursorLocation = adUseClient;
   cObj.CursorLocation = adUseClient;


}

function SetupActiveXObject2()
{
   cObj2 = new ActiveXObject("ADODB.Connection");
   recSet3 = new ActiveXObject("ADODB.Recordset");

   recSet3.CursorLocation = adUseClient;
   cObj2.CursorLocation = adUseClient;

}

function SetupActiveXObject_login()
{
   cObj_login = new ActiveXObject("ADODB.Connection");
   recSet_login = new ActiveXObject("ADODB.Recordset");

   recSet_login.CursorLocation = adUseClient;
   cObj_login.CursorLocation = adUseClient;

}

function SetupActiveXObject_usermanagement()
{
   cObj_usermanagement = new ActiveXObject("ADODB.Connection");
   recSet_usermanagement = new ActiveXObject("ADODB.Recordset");

   recSet_usermanagement.CursorLocation = adUseClient;
   cObj_usermanagement.CursorLocation = adUseClient;

}

//...................................Set Query String...............................................//


function OpenDatabase()
{
   SetupConnectionString()
   SetupActiveXObject()

   cObj.Open(iStr)
      
}
function OpenDatabase2()
{
   SetupConnectionString2()
   SetupActiveXObject2()

   cObj2.Open(iStr2)

}

function OpenDatabase_login()
{
   SetupConnectionString_login()
   SetupActiveXObject_login()

   cObj_login.Open(iStr_login)
}

function OpenDatabase_usermanagement()
{
   SetupConnectionString_usermanagement()
   SetupActiveXObject_usermanagement()

   cObj_usermanagement.Open(iStr_usermanagement)
}

function OpenTable()
{
   recSet.Open(grStr,cObj,adOpenKeyset,adLockOptimistic,adCmdText)

}

function OpenTable2()
{
   recSet2.Open(grStr2,cObj,adOpenKeyset,adLockOptimistic,adCmdText)
}

function OpenTable4()
{
   recSet4.Open(grStr4,cObj,adOpenKeyset,adLockOptimistic,adCmdText)
}

function OpenTable3()
{
   recSet3.Open(grStr3,cObj2,adOpenKeyset,adLockOptimistic,adCmdText)
}

function OpenTable_login()
{
   recSet_login.Open(grStr_login,cObj_login,adOpenKeyset,adLockOptimistic,adCmdText)
}

function OpenTable_usermanagement()
{
   recSet_usermanagement.Open(grStr_usermanagement,cObj_usermanagement,adOpenKeyset,adLockOptimistic,adCmdText)
}
