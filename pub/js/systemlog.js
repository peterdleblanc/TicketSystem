function Database_Class(type)
	{

		this.EmpList=new Array()
		this.EmpName=new Array()
		this.UserCount = 0;
		this.EmpNum = "";
		this.UserAction = "";
		this.SystemName = "";
		this.RequestTime = "";
		this.WorkDone = "";
		this.ActionDate = "";
		this.ActionTime = "";

		this.ReportFrom = "";
		this.ReportTo = "";
		this.ReportName = "";
		this.ReportOn = "";
		this.myDate=new Date();

		//******************************Database Info******************************//
		this.Path = "T:\\All In One\\Reports Viewer\\pub\\databases\\";
		this.iStr = "";
		this.DBName = "";
		this.QryString = "";
		this.adOpenKeyset = '1';
		this.adLockOptimistic = '2';
		this.adCmdText = '1';
		this.adUseClient = '3';
		this.myRow = 1;
		this.myCol = 1;
		this.temp;
		this.temp1;
		this.temp2;
		

		this.cObj = new ActiveXObject("ADODB.Connection");
		this.recSet = new ActiveXObject("ADODB.Recordset");
		
		this.OpenTable = function()
		{
			this.iStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.Path + this.DBName + ";Persist Security Info=False";
			this.cObj.Open(this.iStr)
			this.recSet.Open(this.QryString,this.cObj,this.adOpenKeyset,this.adLockOptimistic,this.adCmdText)
		}
		
		this.CloseTable = function()
		{
			this.recSet.close
			this.cObj.close
			this.recSet = null;
			this.cObj = null;
		}
		
		this.MoveFirst = function()
		{
			if(this.recSet.BOF == 0)
			{
				this.recSet.MoveFirst
			}
		}
		
		this.CreateSpreadSheet = function()
		{
			this.excel = new ActiveXObject ("Excel.Application");
			this.excel.visible = true;
			this.book = this.excel.Workbooks.Add ();
			this.sheet = this.excel.Worksheets(1);
		}
		
	}


