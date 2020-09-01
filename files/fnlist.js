function SetItemValue(ItemName,value)
{
	var obj = document.forms[0].elements[ItemName];
	i = 0;
	err = 1;
	//alert(ItemName);

	while ( (i<obj.length) & (err == 1) )
	{
		if ( obj[i].value == value )
		{
			obj[i].selected = true
			err = 0
		}

		i = i+1
	}

	if ( err == 1 )
	{
		//alert("Sorry! You Set a wrong value in field:"+ItemName)  
	} 
}

function trim(inputString) {
   // Removes leading and trailing spaces from the passed string. Also removes
   // consecutive spaces and replaces it with one space. If something besides
   // a string is passed in (null, custom object, etc.) then return the input.
   if (typeof inputString != "string") { return inputString; }
   var retValue = inputString;
   var ch = retValue.substring(0, 1);
   while (ch == " ") { // Check for spaces at the beginning of the string
      retValue = retValue.substring(1, retValue.length);
      ch = retValue.substring(0, 1);
   }
   ch = retValue.substring(retValue.length-1, retValue.length);
   while (ch == " ") { // Check for spaces at the end of the string
      retValue = retValue.substring(0, retValue.length-1);
      ch = retValue.substring(retValue.length-1, retValue.length);
   }
   while (retValue.indexOf("  ") != -1) { // Note that there are two spaces in the string - look for multiple spaces within the string
      retValue = retValue.substring(0, retValue.indexOf("  ")) + retValue.substring(retValue.indexOf("  ")+1, retValue.length); // Again, there are two spaces in each of the strings
   }
   return retValue; // Return the trimmed string back to the user
} 
// --> Ends the "trim" function

// -- DATE
// ------------------------------------------------------------------
// These functions use the same 'format' strings as the 
// java.text.SimpleDateFormat class, with minor exceptions.
// The format string consists of the following abbreviations:
// 
// Field        | Full Form          | Short Form
// -------------+--------------------+-----------------------
// Year         | yyyy (4 digits)    | yy (2 digits), y (2 or 4 digits)
// Month        | MMM (name or abbr.)| MM (2 digits), M (1 or 2 digits)
// Day of Month | dd (2 digits)      | d (1 or 2 digits)
// Day of Week  | EE (name)          | E (abbr)
// Hour (1-12)  | hh (2 digits)      | h (1 or 2 digits)
// Hour (0-23)  | HH (2 digits)      | H (1 or 2 digits)
// Hour (0-11)  | KK (2 digits)      | K (1 or 2 digits)
// Hour (1-24)  | kk (2 digits)      | k (1 or 2 digits)
// Minute       | mm (2 digits)      | m (1 or 2 digits)
// Second       | ss (2 digits)      | s (1 or 2 digits)
// AM/PM        | a                  |


var MONTH_NAMES=new Array('January','February','March','April','May','June','July','August','September','October','November','December','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec');
var DAY_NAMES=new Array('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sun','Mon','Tue','Wed','Thu','Fri','Sat');
function LZ(x) {return(x<0||x>9?"":"0")+x}

// ------------------------------------------------------------------
// isDate ( date_string, format_string )
// Returns true if date string matches format of format string and
// is a valid date. Else returns false.
// It is recommended that you trim whitespace around the value before
// passing it to this function, as whitespace is NOT ignored!
// ------------------------------------------------------------------
function isDate(val,format) {
	var date=getDateFromFormat(val,format);
	if (date==0) { return false; }
	return true;
	}

// -------------------------------------------------------------------
// compareDates(date1,date1format,date2,date2format)
//   Compare two date strings to see which is greater.
//   Returns:
//   1 if date1 is greater than date2
//   0 if date2 is greater than date1 of if they are the same
//  -1 if either of the dates is in an invalid format
// -------------------------------------------------------------------
function compareDates(date1,dateformat1,date2,dateformat2) {
	var d1=getDateFromFormat(date1,dateformat1);
	var d2=getDateFromFormat(date2,dateformat2);
	if (d1==0 || d2==0) {
		return -1;
		}
	else if (d1 > d2) {
		return 1;
		}
	return 0;
	}

// ------------------------------------------------------------------
// formatDate (date_object, format)
// Returns a date in the output format specified.
// The format string uses the same abbreviations as in getDateFromFormat()
// ------------------------------------------------------------------
function formatDate(date,format) {
	format=format+"";
	var result="";
	var i_format=0;
	var c="";
	var token="";
	var y=date.getYear()+"";
	var M=date.getMonth()+1;
	var d=date.getDate();
	var E=date.getDay();
	var H=date.getHours();
	var m=date.getMinutes();
	var s=date.getSeconds();
	var yyyy,yy,MMM,MM,dd,hh,h,mm,ss,ampm,HH,H,KK,K,kk,k;
	// Convert real date parts into formatted versions
	var value=new Object();
	if (y.length < 4) {y=""+(y-0+1900);}
	value["y"]=""+y;
	value["yyyy"]=y;
	value["yy"]=y.substring(2,4);
	value["M"]=M;
	value["MM"]=LZ(M);
	value["MMM"]=MONTH_NAMES[M-1];
	value["d"]=d;
	value["dd"]=LZ(d);
	value["E"]=DAY_NAMES[E+7];
	value["EE"]=DAY_NAMES[E];
	value["H"]=H;
	value["HH"]=LZ(H);
	if (H==0){value["h"]=12;}
	else if (H>12){value["h"]=H-12;}
	else {value["h"]=H;}
	value["hh"]=LZ(value["h"]);
	if (H>11){value["K"]=H-12;} else {value["K"]=H;}
	value["k"]=H+1;
	value["KK"]=LZ(value["K"]);
	value["kk"]=LZ(value["k"]);
	if (H > 11) { value["a"]="PM"; }
	else { value["a"]="AM"; }
	value["m"]=m;
	value["mm"]=LZ(m);
	value["s"]=s;
	value["ss"]=LZ(s);
	while (i_format < format.length) {
		c=format.charAt(i_format);
		token="";
		while ((format.charAt(i_format)==c) && (i_format < format.length)) {
			token += format.charAt(i_format++);
			}
		if (value[token] != null) { result=result + value[token]; }
		else { result=result + token; }
		}
	return result;
	}
	
// ------------------------------------------------------------------
// Utility functions for parsing in getDateFromFormat()
// ------------------------------------------------------------------
function isInteger(val) {
	var digits="1234567890";
	for (var i=0; i < val.length; i++) {
		if (digits.indexOf(val.charAt(i))==-1) { return false; }
		}
	return true;
	}
function _getInt(str,i,minlength,maxlength) {
	for (var x=maxlength; x>=minlength; x--) {
		var token=str.substring(i,i+x);
		if (token.length < minlength) { return null; }
		if (isInteger(token)) { return token; }
		}
	return null;
	}
	
// ------------------------------------------------------------------
// getDateFromFormat( date_string , format_string )
//
// This function takes a date string and a format string. It matches
// If the date string matches the format string, it returns the 
// getTime() of the date. If it does not match, it returns 0.
// ------------------------------------------------------------------
function getDateFromFormat(val,format) {
	val=val+"";
	format=format+"";
	var i_val=0;
	var i_format=0;
	var c="";
	var token="";
	var token2="";
	var x,y;
	var now=new Date();
	var year=now.getYear();
	var month=now.getMonth()+1;
	var date=1;
	var hh=now.getHours();
	var mm=now.getMinutes();
	var ss=now.getSeconds();
	var ampm="";
	
	while (i_format < format.length) {
		// Get next token from format string
		c=format.charAt(i_format);
		token="";
		while ((format.charAt(i_format)==c) && (i_format < format.length)) {
			token += format.charAt(i_format++);
			}
		// Extract contents of value based on format token
		if (token=="yyyy" || token=="yy" || token=="y") {
			if (token=="yyyy") { x=4;y=4; }
			if (token=="yy")   { x=2;y=2; }
			if (token=="y")    { x=2;y=4; }
			year=_getInt(val,i_val,x,y);
			if (year==null) { return 0; }
			i_val += year.length;
			if (year.length==2) {
				if (year > 70) { year=1900+(year-0); }
				else { year=2000+(year-0); }
				}
			}
		else if (token=="MMM"){
			month=0;
			for (var i=0; i<MONTH_NAMES.length; i++) {
				var month_name=MONTH_NAMES[i];
				if (val.substring(i_val,i_val+month_name.length).toLowerCase()==month_name.toLowerCase()) {
					month=i+1;
					if (month>12) { month -= 12; }
					i_val += month_name.length;
					break;
					}
				}
			if ((month < 1)||(month>12)){return 0;}
			}
		else if (token=="EE"||token=="E"){
			for (var i=0; i<DAY_NAMES.length; i++) {
				var day_name=DAY_NAMES[i];
				if (val.substring(i_val,i_val+day_name.length).toLowerCase()==day_name.toLowerCase()) {
					i_val += day_name.length;
					break;
					}
				}
			}
		else if (token=="MM"||token=="M") {
			month=_getInt(val,i_val,token.length,2);
			if(month==null||(month<1)||(month>12)){return 0;}
			i_val+=month.length;}
		else if (token=="dd"||token=="d") {
			date=_getInt(val,i_val,token.length,2);
			if(date==null||(date<1)||(date>31)){return 0;}
			i_val+=date.length;}
		else if (token=="hh"||token=="h") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<1)||(hh>12)){return 0;}
			i_val+=hh.length;}
		else if (token=="HH"||token=="H") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<0)||(hh>23)){return 0;}
			i_val+=hh.length;}
		else if (token=="KK"||token=="K") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<0)||(hh>11)){return 0;}
			i_val+=hh.length;}
		else if (token=="kk"||token=="k") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<1)||(hh>24)){return 0;}
			i_val+=hh.length;hh--;}
		else if (token=="mm"||token=="m") {
			mm=_getInt(val,i_val,token.length,2);
			if(mm==null||(mm<0)||(mm>59)){return 0;}
			i_val+=mm.length;}
		else if (token=="ss"||token=="s") {
			ss=_getInt(val,i_val,token.length,2);
			if(ss==null||(ss<0)||(ss>59)){return 0;}
			i_val+=ss.length;}
		else if (token=="a") {
			if (val.substring(i_val,i_val+2).toLowerCase()=="am") {ampm="AM";}
			else if (val.substring(i_val,i_val+2).toLowerCase()=="pm") {ampm="PM";}
			else {return 0;}
			i_val+=2;}
		else {
			if (val.substring(i_val,i_val+token.length)!=token) {return 0;}
			else {i_val+=token.length;}
			}
		}
	// If there are any trailing characters left in the value, it doesn't match
	if (i_val != val.length) { return 0; }
	// Is date valid for month?
	if (month==2) {
		// Check for leap year
		if ( ( (year%4==0)&&(year%100 != 0) ) || (year%400==0) ) { // leap year
			if (date > 29){ return false; }
			}
		else { if (date > 28) { return false; } }
		}
	if ((month==4)||(month==6)||(month==9)||(month==11)) {
		if (date > 30) { return false; }
		}
	// Correct hours value
	if (hh<12 && ampm=="PM") { hh=hh-0+12; }
	else if (hh>11 && ampm=="AM") { hh-=12; }
	var newdate=new Date(year,month-1,date,hh,mm,ss);
	return newdate.getTime();
	}

function IsValidNo(field, Value)
{
	if (isNaN(Value))
	{
		alert("Please enter a valid number.");
		field.select();
		field.focus();
	}	
}
function RoundOff(value, precision)
{
        value = "" + value //convert value to string
        precision = parseInt(precision);

        var whole = "" + Math.round(value * Math.pow(10, precision));

        var decPoint = whole.length - precision;

        if(decPoint != 0)
        {
                result = whole.substring(0, decPoint);
                result += ".";
                result += whole.substring(decPoint, whole.length);
        }
        else
        {
                result = "0." + whole;
        }
        return result;
}
	

function validateReq(formField,fieldLabel)
{
  var result = true;
  if (formField.value == "")
  {
    alert('Please enter a value for the "' + fieldLabel +'" field.');
    formField.focus();
    result = false;
  }
 
  return result;
}
function validateReqNo(formField,fieldLabel)
{
  var result = true;
  if (formField.value == "" ||  formField.value==0)  	
  	{
	    alert('Please enter a value for the "' + fieldLabel +'" field.');
	    formField.focus();
	    result = false;
  	}  
 
  return result;
}
function validateReq2(formField,fieldLabel,defaultField,defaultLabel)
{
  var result = true;
  if (formField.value == "")
  {
  	if (defaultField.value == "" )
  	{
  		alert('Please enter a value for the "' + defaultLabel +'" field.');
  		defaultField.focus();
  		result = false;
  	}
  	else
  	{
	    alert('Please enter a value for the "' + fieldLabel +'" field.');
	    formField.focus();
	    result = false;
  	}
  }
 
  return result;
}
function validateReqNo2(formField,fieldLabel,defaultField,defaultLabel)
{
  var result = true;
  if (formField.value == "")
  {
  	if (defaultField.value == "" &&  fieldLabel.value!=0)
  	{
  		alert('Please enter a value for the "' + defaultLabel +'" field.');
  		defaultField.focus();
  		result = false;
  	}
  	else
  	{
	    alert('Please enter a value for the "' + fieldLabel +'" field.');
	    formField.focus();
	    result = false;
  	}
  }
 
  return result;
}

function blankReq(formField,fieldLabel)
{
  var result = true;
  if (formField.value != "")
    result = false;
 
  return result;
}

//*************** Start Request Details ****************//
function formRequestDetails(theForm)
{
	{
	if(!validateReq(theForm.requestor,"Requestor"))
		return false;
	if(!validateReq(theForm.reqemail,"Email"))
		return false;
	if(!validateReq(theForm.reqmgr,"Requestor's Manager"))
		return false;
	if(!validateReq(theForm.reqmgremail,"Manager Email"))
		return false;
	if(!validateReq(theForm.customer,"Customer"))
		return false;
	if(!validateReq(theForm.appl,"Application"))
		return false; 
	if(!validateReq(theForm.priority,"Priority"))
		return false;		
	if(!validateReq(theForm.reqtitle,"Request Title"))
		return false;
	if(!validateReq(theForm.Desc,"Description"))
		return false;
	if(!validateReq(theForm.expclosedate,"Expected Go Live Date"))
		return false; 
	}
	return true;
}
//*************** End Request Details ****************//

//*************** Start IT QC Details ****************//
function formITQCDetails(theForm)
{
	if(theForm.TeamLead.value=="" && theForm.ExpQC.value=="" )
		return true;
	else
	{	
		{
			if(!validateReq(theForm.TeamLead,"TeamLead"))
				return false;
		}
	}
	return true;
}		
//*************** END IT QC Details ****************//

//*************** Start IT Details ****************//
function formITDetails(theForm)
{
	if(theForm.Developer.value=="" && (theForm.EstManHour.value=="" || theForm.EstManHour.value=="0" ) && (theForm.EstTotalCost.value=="" || theForm.EstTotalCost.value=="0") && theForm.ExpStartDate.value=="" && theForm.ExpEndDate.value==""  && theForm.RemarksIT.value=="" )
		return true;
	else
	{	
		{
			if(!validateReq(theForm.TeamLead,"TeamLead"))
				return false;
			if(!validateReq(theForm.Developer,"Developer"))
				return false;
			if(!validateReq(theForm.ExpQC,"Exp QC"))
				return false;
			if(!validateReqNo(theForm.EstManHour,"Estimated ManHour"))
				return false;
			if(!validateReq(theForm.EstTotalCost,"Estimated Total Cost"))
				return false;
			if(!validateReq(theForm.ExpStartDate,"Exp. Start Date"))
				return false;
			if(!validateReq(theForm.ExpEndDate,"Exp. Complete Date"))
				return false; 
//			if(!validateReq(theForm.RemarksIT,"Remarks"))
//				return false;		
		}
		return true;
	}
}
//*************** End IT Details ****************//


//*************** Start Programmer's Update ****************//

function formProgStart(theForm)
{
	if(theForm.ActStartDate.value=="" && theForm.RemarksDev.value=="" && (theForm.ActManHour.value=="" || theForm.ActManHour.value=="0") && (theForm.ActTotalCost.value=="" || theForm.ActTotalCost.value=="0") && theForm.ActEndDate.value=="" )
		return true;
	else	
	{
		if(!validateReq(theForm.TeamLead,"TeamLead (prog)"))
			return false;
	}
	return true;
}

function formProgDetails(theForm)
{
	if(theForm.ActStartDate.value=="" && theForm.RemarksDev.value=="" && (theForm.ActManHour.value=="" || theForm.ActManHour.value=="0") && (theForm.ActTotalCost.value=="" || theForm.ActTotalCost.value=="0") && theForm.ActEndDate.value=="" )
		return true;
	else	
	{
		if(!validateReq(theForm.ActStartDate,"Act. Start Date"))
			return false;
		
		if (theForm.ActEndDate.value!="")
		{
			if(!validateReqNo(theForm.ActManHour,"Act. ManHour"))
				return false;
		}
	}
	return true;
}

//*************** End Programmer's Update ****************//

//*************** Start QC's Update ****************//
// Start Checking for empty fields //

function formQCDetails(theForm)
{
	if(theForm.ActQC.value=="" && theForm.StatusQC.value=="" && theForm.RemarksQC.value=="" && theForm.UATReadyDate.value=="")
		return true;
	else	
	{
		if(!validateReq(theForm.ActStartDate,"Actual Start Date"))
			return false;		
		if(!validateReq(theForm.ActEndDate,"Act. Completed Date"))
			return false;
				
		if(!validateReq(theForm.ActQC,"Actual QC"))
			return false;
		if(theForm.UATFailed.value!="Y")
		{	
			if(!validateReq(theForm.StatusQC,"Status QC"))
				return false;
		}
//		if(!validateReq(theForm.RemarksQC,"Remarks QC"))
//			return false;

		if(theForm.StatusQC.value=="P")
		{
			if(!validateReq(theForm.UATReadyDate,"UAT Ready Date"))
				return false;
		}
	}
	return true;
}
//*************** End QC's Update ****************//


//*************** Start UAT Details ****************//
// Start Checking for empty fields //

function formUatDetails(theForm)
{
	if(theForm.UserUAT.value=="" && theForm.StatusUAT.value=="" && theForm.RemarksUAT.value=="" && theForm.ExpCutinDate.value=="")
		return true;
	else
	{
		// added by yusming on 18 Dec 2004
		if ((theForm.UserUAT.value == theForm.appmgr.value) || (theForm.UserUAT.value == theForm.TeamLead.value) || (theForm.UserUAT.value == theForm.Developer.value) || (theForm.UserUAT.value == theForm.ActQC.value))
		{
			alert("You are not authorised to perform this function.\n  Pls kindly inform Requestor to perform UAT");
			theForm.UserUAT.focus();
			return false;
		}
		if((theForm.UATFailed.value=="Y") && (theForm.StatusQC.value=="P"))
		{
			alert("Please clear the UAT Details (You cannot key in the UAT details for UAT Failed Status)");
			theForm.UserUAT.focus();
			return false;
		}
		if(!validateReq(theForm.ActQC,"Actual QC"))
			return false;
		if(!validateReq(theForm.UATReadyDate,"UAT Ready Date"))
			return false;	
		if(!validateReq(theForm.UserUAT,"User UAT"))
			return false;
		if(!validateReq(theForm.StatusUAT,"Status UAT"))
			return false;
//		if(!validateReq(theForm.RemarksUAT,"Remarks UAT"))
//			return false;
		if(theForm.StatusUAT.value=="P")
			{
			if(!validateReq(theForm.ExpCutinDate,"Exp Cut-in Date"))
				return false;
			}
	}
	return true;
}
//*************** End UAT Details ****************//

//*************** Start Deploy Details ****************//

function formDeployDetails(theForm)
{
	if(theForm.DeployUser.value=="" && theForm.ActCutinDate.value=="" && theForm.RemarksDeploy.value=="")
		return true;
	else	
	{
		if(!validateReq(theForm.UserUAT,"User UAT"))
			return false;
		if(!validateReq(theForm.ExpCutinDate,"Exp Cut-in Date"))
			return false;
		if(!validateReq(theForm.DeployUser,"Deploy User"))
			return false;
		if(!validateReq(theForm.ActCutinDate,"Act Cut-in Date"))
			return false;
//		if(!validateReq(theForm.RemarksDeploy,"Remarks Deploy"))
//			return false;
			
	}
	return true;
}
//*************** End Deploy Details****************//

//email validation start
function isEmailAddr(email)
{
  var result = false
  var theStr = new String(email)
  var index = theStr.indexOf("@");
  if (index > 0)
  {
    var pindex = theStr.indexOf(".",index);
    if ((pindex > index+1) && (theStr.length > pindex+1))
	result = true;
  }
  return result;
}

function EmailValidator(theForm)
{

  if (theForm.reqemail.value == "")
  {
    alert("Please enter a value for the \"email\" field.");
    theForm.reqemail.focus();
    return (false);
  }

  if (!isEmailAddr(theForm.reqemail.value))
  {
    alert("Please enter a complete email address in the form: yourname@yourdomain.com");
    theForm.reqemail.focus();
    return (false);
  }
   
  if (theForm.reqemail.value.length < 6)
  {
    alert("Please enter at least 6 characters in the \"email\" field.");
    theForm.reqemail.focus();
    return (false);
  }    if (theForm.reqmgremail.value == "")
  {
    alert("Please enter a value for the \"email\" field.");
    theForm.reqmgremail.focus();
    return (false);
  }

  if (!isEmailAddr(theForm.reqmgremail.value))
  {
    alert("Please enter a complete email address in the form: yourname@yourdomain.com");
    theForm.reqmgremail.focus();
    return (false);
  }
   
  if (theForm.reqmgremail.value.length < 6)
  {
    alert("Please enter at least 6 characters in the \"email\" field.");
    theForm.reqmgremail.focus();
    return (false);
  }
  return (true);
}


//email validation end
//--------------------------------------------------------------------------------------------------//


// ------------------------------------------------------------------
// These functions use the same 'format' strings as the 
// java.text.SimpleDateFormat class, with minor exceptions.
// The format string consists of the following abbreviations:
// 
// Field        | Full Form          | Short Form
// -------------+--------------------+-----------------------
// Year         | yyyy (4 digits)    | yy (2 digits), y (2 or 4 digits)
// Month        | MMM (name or abbr.)| MM (2 digits), M (1 or 2 digits)
// Day of Month | dd (2 digits)      | d (1 or 2 digits)
// Day of Week  | EE (name)          | E (abbr)
// Hour (1-12)  | hh (2 digits)      | h (1 or 2 digits)
// Hour (0-23)  | HH (2 digits)      | H (1 or 2 digits)
// Hour (0-11)  | KK (2 digits)      | K (1 or 2 digits)
// Hour (1-24)  | kk (2 digits)      | k (1 or 2 digits)
// Minute       | mm (2 digits)      | m (1 or 2 digits)
// Second       | ss (2 digits)      | s (1 or 2 digits)
// AM/PM        | a                  |


var MONTH_NAMES=new Array('January','February','March','April','May','June','July','August','September','October','November','December','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec');
var DAY_NAMES=new Array('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sun','Mon','Tue','Wed','Thu','Fri','Sat');
function LZ(x) {return(x<0||x>9?"":"0")+x}

// ------------------------------------------------------------------
// isDate ( date_string, format_string )
// Returns true if date string matches format of format string and
// is a valid date. Else returns false.
// It is recommended that you trim whitespace around the value before
// passing it to this function, as whitespace is NOT ignored!
// ------------------------------------------------------------------
function isDate(val,format) {
	var date=getDateFromFormat(val,format);
	if (date==0) { return false; }
	return true;
	}

// -------------------------------------------------------------------
// compareDates(date1,date1format,date2,date2format)
//   Compare two date strings to see which is greater.
//   Returns:
//   1 if date1 is greater than date2
//   0 if date2 is greater than date1 of if they are the same
//  -1 if either of the dates is in an invalid format
// -------------------------------------------------------------------
function compareDates(date1,dateformat1,date2,dateformat2) {
	var d1=getDateFromFormat(date1,dateformat1);
	var d2=getDateFromFormat(date2,dateformat2);
	if (d1==0 || d2==0) {
		return -1;
		}
	else if (d1 > d2) {
		return 1;
		}
	return 0;
	}

// ------------------------------------------------------------------
// formatDate (date_object, format)
// Returns a date in the output format specified.
// The format string uses the same abbreviations as in getDateFromFormat()
// ------------------------------------------------------------------
function formatDate(date,format) {
	format=format+"";
	var result="";
	var i_format=0;
	var c="";
	var token="";
	var y=date.getYear()+"";
	var M=date.getMonth()+1;
	var d=date.getDate();
	var E=date.getDay();
	var H=date.getHours();
	var m=date.getMinutes();
	var s=date.getSeconds();
	var yyyy,yy,MMM,MM,dd,hh,h,mm,ss,ampm,HH,H,KK,K,kk,k;
	// Convert real date parts into formatted versions
	var value=new Object();
	if (y.length < 4) {y=""+(y-0+1900);}
	value["y"]=""+y;
	value["yyyy"]=y;
	value["yy"]=y.substring(2,4);
	value["M"]=M;
	value["MM"]=LZ(M);
	value["MMM"]=MONTH_NAMES[M-1];
	value["d"]=d;
	value["dd"]=LZ(d);
	value["E"]=DAY_NAMES[E+7];
	value["EE"]=DAY_NAMES[E];
	value["H"]=H;
	value["HH"]=LZ(H);
	if (H==0){value["h"]=12;}
	else if (H>12){value["h"]=H-12;}
	else {value["h"]=H;}
	value["hh"]=LZ(value["h"]);
	if (H>11){value["K"]=H-12;} else {value["K"]=H;}
	value["k"]=H+1;
	value["KK"]=LZ(value["K"]);
	value["kk"]=LZ(value["k"]);
	if (H > 11) { value["a"]="PM"; }
	else { value["a"]="AM"; }
	value["m"]=m;
	value["mm"]=LZ(m);
	value["s"]=s;
	value["ss"]=LZ(s);
	while (i_format < format.length) {
		c=format.charAt(i_format);
		token="";
		while ((format.charAt(i_format)==c) && (i_format < format.length)) {
			token += format.charAt(i_format++);
			}
		if (value[token] != null) { result=result + value[token]; }
		else { result=result + token; }
		}
	return result;
	}
	
// ------------------------------------------------------------------
// Utility functions for parsing in getDateFromFormat()
// ------------------------------------------------------------------
function isInteger(val) {
	var digits="1234567890";
	for (var i=0; i < val.length; i++) {
		if (digits.indexOf(val.charAt(i))==-1) { return false; }
		}
	return true;
	}
function _getInt(str,i,minlength,maxlength) {
	for (var x=maxlength; x>=minlength; x--) {
		var token=str.substring(i,i+x);
		if (token.length < minlength) { return null; }
		if (isInteger(token)) { return token; }
		}
	return null;
	}
	
// ------------------------------------------------------------------
// getDateFromFormat( date_string , format_string )
//
// This function takes a date string and a format string. It matches
// If the date string matches the format string, it returns the 
// getTime() of the date. If it does not match, it returns 0.
// ------------------------------------------------------------------
function getDateFromFormat(val,format) {
	val=val+"";
	format=format+"";
	var i_val=0;
	var i_format=0;
	var c="";
	var token="";
	var token2="";
	var x,y;
	var now=new Date();
	var year=now.getYear();
	var month=now.getMonth()+1;
	var date=1;
	var hh=now.getHours();
	var mm=now.getMinutes();
	var ss=now.getSeconds();
	var ampm="";
	
	while (i_format < format.length) {
		// Get next token from format string
		c=format.charAt(i_format);
		token="";
		while ((format.charAt(i_format)==c) && (i_format < format.length)) {
			token += format.charAt(i_format++);
			}
		// Extract contents of value based on format token
		if (token=="yyyy" || token=="yy" || token=="y") {
			if (token=="yyyy") { x=4;y=4; }
			if (token=="yy")   { x=2;y=2; }
			if (token=="y")    { x=2;y=4; }
			year=_getInt(val,i_val,x,y);
			if (year==null) { return 0; }
			i_val += year.length;
			if (year.length==2) {
				if (year > 70) { year=1900+(year-0); }
				else { year=2000+(year-0); }
				}
			}
		else if (token=="MMM"){
			month=0;
			for (var i=0; i<MONTH_NAMES.length; i++) {
				var month_name=MONTH_NAMES[i];
				if (val.substring(i_val,i_val+month_name.length).toLowerCase()==month_name.toLowerCase()) {
					month=i+1;
					if (month>12) { month -= 12; }
					i_val += month_name.length;
					break;
					}
				}
			if ((month < 1)||(month>12)){return 0;}
			}
		else if (token=="EE"||token=="E"){
			for (var i=0; i<DAY_NAMES.length; i++) {
				var day_name=DAY_NAMES[i];
				if (val.substring(i_val,i_val+day_name.length).toLowerCase()==day_name.toLowerCase()) {
					i_val += day_name.length;
					break;
					}
				}
			}
		else if (token=="MM"||token=="M") {
			month=_getInt(val,i_val,token.length,2);
			if(month==null||(month<1)||(month>12)){return 0;}
			i_val+=month.length;}
		else if (token=="dd"||token=="d") {
			date=_getInt(val,i_val,token.length,2);
			if(date==null||(date<1)||(date>31)){return 0;}
			i_val+=date.length;}
		else if (token=="hh"||token=="h") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<1)||(hh>12)){return 0;}
			i_val+=hh.length;}
		else if (token=="HH"||token=="H") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<0)||(hh>23)){return 0;}
			i_val+=hh.length;}
		else if (token=="KK"||token=="K") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<0)||(hh>11)){return 0;}
			i_val+=hh.length;}
		else if (token=="kk"||token=="k") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<1)||(hh>24)){return 0;}
			i_val+=hh.length;hh--;}
		else if (token=="mm"||token=="m") {
			mm=_getInt(val,i_val,token.length,2);
			if(mm==null||(mm<0)||(mm>59)){return 0;}
			i_val+=mm.length;}
		else if (token=="ss"||token=="s") {
			ss=_getInt(val,i_val,token.length,2);
			if(ss==null||(ss<0)||(ss>59)){return 0;}
			i_val+=ss.length;}
		else if (token=="a") {
			if (val.substring(i_val,i_val+2).toLowerCase()=="am") {ampm="AM";}
			else if (val.substring(i_val,i_val+2).toLowerCase()=="pm") {ampm="PM";}
			else {return 0;}
			i_val+=2;}
		else {
			if (val.substring(i_val,i_val+token.length)!=token) {return 0;}
			else {i_val+=token.length;}
			}
		}
	// If there are any trailing characters left in the value, it doesn't match
	if (i_val != val.length) { return 0; }
	// Is date valid for month?
	if (month==2) {
		// Check for leap year
		if ( ( (year%4==0)&&(year%100 != 0) ) || (year%400==0) ) { // leap year
			if (date > 29){ return false; }
			}
		else { if (date > 28) { return false; } }
		}
	if ((month==4)||(month==6)||(month==9)||(month==11)) {
		if (date > 30) { return false; }
		}
	// Correct hours value
	if (hh<12 && ampm=="PM") { hh=hh-0+12; }
	else if (hh>11 && ampm=="AM") { hh-=12; }
	var newdate=new Date(year,month-1,date,hh,mm,ss);
	return newdate.getTime();
	}

function IsValidNo(field, Value)
{
	if (isNaN(Value))
	{
		alert("Please enter a valid number.");
		field.select();
		field.focus();
	}	
}
function RoundOff(value, precision)
{
        value = "" + value //convert value to string
        precision = parseInt(precision);

        var whole = "" + Math.round(value * Math.pow(10, precision));

        var decPoint = whole.length - precision;

        if(decPoint != 0)
        {
                result = whole.substring(0, decPoint);
                result += ".";
                result += whole.substring(decPoint, whole.length);
        }
        else
        {
                result = "0." + whole;
        }
        return result;
}
	