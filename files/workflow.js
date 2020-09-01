/******* QC Fields WorkFlow ********/

/******* Programmer Fields WorkFlow ********/
function ProgblankReq(formField,fieldLabel,prevField1,prevField2,prevField3,prevField4)
{
  if (formField.value != "")
  {
  	if(prevField1.value == "")
	{
		alert('You have to complete the IT Details');
		prevField1.focus(); 
	    return false;
	}
	else if(prevField2.value =="")
	{
		alert('You have to complete the IT Details');
		prevField2.focus(); 
	    return false;
	}
 	else if(prevField3.value == "")
	{
		alert('You have to complete the IT Details');
		prevField3.focus(); 
	    return false;
	}
	else if(prevField4.value =="")
	{
		alert('You have to complete the IT Details');
		prevField4.focus(); 
	    return false;
	}

  }
  return true;
}

function formProgBlank(theForm)
{
	if(!ProgblankReq(theForm.ActStartDate,"Actual Start Date", theForm.TeamLead, theForm.Developer, theForm.ExpStartDate, theForm.ExpEndDate))
		return false;
	if(!ProgblankReq(theForm.ActEndDate,"Actual End Date", theForm.TeamLead, theForm.Developer, theForm.ExpStartDate, theForm.ExpEndDate))
		return false;
		
	return true;
}

/******* QC Fields WorkFlow ********/

function QCblankReq(formField,fieldLabel,prevField1,prevField2)
{
  if (formField.value != "")
  {
  	if(prevField1.value == "")
	{
		alert('You have to complete the Programmer Update');
		prevField1.focus(); 
	    return false;
	}
	else if(prevField2.value =="")
	{
		alert('You have to complete the Programmer Update');
		prevField2.focus(); 
	    return false;
	}

  }
  return true;
}

function formQCBlank(theForm)
{
	if(!QCblankReq(theForm.ActQC,"Actual QC",theForm.ActStartDate, theForm.ActEndDate))
		return false;
	if(!QCblankReq(theForm.StatusQC,"Status QC",theForm.ActStartDate, theForm.ActEndDate))
		return false;
	if(!QCblankReq(theForm.RemarksQC,"Remarks UAT",theForm.ActStartDate, theForm.ActEndDate))
		return false;	
	if(!QCblankReq(theForm.UATReadyDate,"Cutting Date",theForm.ActStartDate, theForm.ActEndDate))
		return false;	
		
	return true;
}

/******* UAT Fields WorkFlow ********/
function uatblankReq(formField,fieldLabel,prevField1)
{
  if (formField.value != "")
  {
  	if(prevField1.value == "")
	{
		alert('You have to complete the QC Details');
		prevField1.focus(); 
	    return false;
	}
  }
  return true;
}

function formUATBlank(theForm)
{
	if(!uatblankReq(theForm.UserUAT,"UAT User",theForm.UATReadyDate))
		return false;
	if(!uatblankReq(theForm.StatusUAT,"Status UAT",theForm.UATReadyDate))
		return false;
	if(!uatblankReq(theForm.RemarksUAT,"Remarks UAT",theForm.UATReadyDate))
		return false;	
	if(!uatblankReq(theForm.ExpCutinDate,"Cutting Date",theForm.UATReadyDate))
		return false;	
		
	return true;
}
/******* Deploy Fields WorkFlow ********/
function blankReq(formField,fieldLabel,prevField1,prevField2,prevField3)
{
  if (formField.value != "")
  {
  	if(prevField1.value == "")
	{
		alert('You have to complete the UAT Details');
		prevField1.focus(); 
	    return false;
	}
	else if(prevField2.value =="")
	{
		alert('You have to complete the UAT Details');
		prevField2.focus(); 
	    return false;
	}
	else if(prevField3.value == "")
	{
		alert('You have to complete the UAT Details');
		prevField3.focus(); 
	    return false;		
	}

  }
  return true;
}

function formDeployBlank(theForm)
{
	if(!blankReq(theForm.DeployUser,"Deploy User",theForm.UserUAT,theForm.StatusUAT,theForm.ExpCutinDate))
		return false;
	if(!blankReq(theForm.ActCutinDate,"Act Cut-in Date",theForm.UserUAT,theForm.StatusUAT,theForm.ExpCutinDate))
		return false;
	if(!blankReq(theForm.RemarksDeploy,"Remarks Deploy",theForm.UserUAT,theForm.StatusUAT,theForm.ExpCutinDate))
		return false;	
	
	return true;
}
//*************** End Deploy Details****************//