	function CheckLogin_(name)
	{
		var LoginStatus = "";
		LoginStatus = autologin_systemclass.verify(name)
		
		if (LoginStatus == 1)
		{
			CheckLogin(autologin_systemclass.readPointer('Username'),autologin_systemclass.readPointer('Password'))
		}
	}
	
	function AutoLogin(system)
	{
		this.systems = Array(
			'launcher',
			'default_tickets',
			'review_tickets',
			'trucks',
			'trucks_callbacks',
			'reports_viewer',
			'time_adjuster'
		);
		
		this.date = new Date();
		this.expires = "";
		
		this.createPointer = function(name,value,days)
		{
			for(var i=0;i < this.systems.length;i++)
			{
				if (this.systems[i] == system)
				{
					if (days) {
						this.date.setTime(this.date.getTime()+(days*24*60*60*1000));
						this.expires = "; expires="+this.date.toGMTString();
					}
					else this.expires = "";
					
					document.cookie = name+"="+value+this.expires+"; path=/";
				}
			}
		}
		
		this.nameEQ = "";
		this.ca = "";
		this.c = 0;
		
		this.readPointer = function(name)
		{
			this.nameEQ = name + "=";
			this.ca = document.cookie.split(';');
			
			for(var i=0;i < this.ca.length;i++)
			{
				this.c = this.ca[i];
				
				while (this.c.charAt(0)==' ') this.c = this.c.substring(1,this.c.length);
				{
					if (this.c.indexOf(this.nameEQ) == 0)
					{
						return this.c.substring(this.nameEQ.length,this.c.length);
					}
				}
			}
			return null;
		}
		
		this.erasePointer = function(name)
		{
			this.createPointer(name,"",-1);
		}
		
		this.ReturnedValue = "";
		
		this.verify = function(name)
		{
			this.ReturnedValue = this.readPointer(name)
			return this.ReturnedValue;
		}
	}