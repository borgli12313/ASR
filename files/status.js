/******* Check if the status if fail, then prohibit to proceed ********/


/******* The Deploy fields must be blank if the UAT status is fail **************/
function DeployMustBlank(theForm)
{
	//Deploy must be blank if the qc is fail
	if(theForm.DeployUser.value !="")
	{
		alert('You can not proceed  to Deploy if the UAT status is fail');
		theForm.DeployUser.focus(); 
	    return false;
	}
	if(theForm.ActCutinDate.value !="")
	{
		alert('You can not proceed to Deploy if the UAT status is fail');
		theForm.ActCutinDate.focus(); 
	    return false;
	}
	if(theForm.RemarksDeploy.value !="")
	{
		alert('You can not proceed  to Deploy if the UAT status is fail');
		theForm.RemarksDeploy.focus(); 
	    return false;
	}
	return true;
}

/******* The UAT fields must be blank if the QC status is fail **************/
function UATMustBlank(theForm)
{
	//the uat section must be blank if the qc is fail
	if(theForm.UserUAT.value !="")
	{
		alert('You can not proceed to UAT if the QC status is fail');
		theForm.UserUAT.focus(); 
	    return false;
	}
	if(theForm.StatusUAT.value !="")
	{
		alert('You can not proceed to UAT if the QC status is fail');
		theForm.StatusUAT.focus(); 
	    return false;
	}
	if(theForm.RemarksUAT.value !="")
	{
		alert('You can not proceed to UAT if the QC status is fail');
		theForm.RemarksUAT.focus(); 
	    return false;
	}
	if(theForm.ExpCutinDate.value !="")
	{
		alert('You can not proceed to UAT if the QC status is fail');
		theForm.ExpCutinDate.focus(); 
	    return false;
	}
	
	//Deploy section has to be empty also if the qc is fail
	if(theForm.DeployUser.value !="")
	{
		alert('You can not proceed  to Deploy if the QC status is fail');
		theForm.DeployUser.focus(); 
	    return false;
	}
	if(theForm.ActCutinDate.value !="")
	{
		alert('You can not proceed to Deploy if the QC status is fail');
		theForm.ActCutinDate.focus(); 
	    return false;
	}
	if(theForm.RemarksDeploy.value !="")
	{
		alert('You can not proceed  to Deploy if the QC status is fail');
		theForm.RemarksDeploy.focus(); 
	    return false;
	}
	return true;
}

function StatusCheck(theForm)
{
	if(theForm.StatusUAT.value =="F")
	{
		if(!DeployMustBlank(theForm))
			return false;
	}
	
	if(theForm.StatusQC.value =="F")
	{
		if(!UATMustBlank(theForm))
			return false;
	}
	return true;
}
/******** End Status Validation**********/