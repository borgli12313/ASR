<script Language="JavaScript">
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
</script>