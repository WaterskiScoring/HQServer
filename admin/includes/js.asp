<script language="JavaScript">
<!--  //Begin

	function goToURL(url) 
	{
		window.location = url;
	}
	
	function NewWindow(mypage, myname, w, h, scroll)
	{
		winprops = 'height='+h+',width='+w+',scrollbars='+scroll+''
		win = window.open(mypage, myname, winprops)
		if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
	}
	
		function checkPw(form) 
	{
		pw1 = form.Password.value;
		pw2 = form.Password2.value;
	
		if (pw1 != pw2) 
		{
			alert ("\nYou did not enter the same new password twice. Please re-enter your password.")
			return false;
		}
		else return true;
	}
	
	function confirmDelete(type, name, form) 
	{
		if (confirm("Are you sure you want to PERMENANTLY delete the " + type + " " + name + "?")) {
		return true;
	}
	else {
		return false;
		}
	}

	function confirmDeleteText(text, form) 
	{
		if (confirm(text)) {
		return true;
	}
	else {
		return false;
		}
	}
	
	function checkChoices(numChoices)
	{
		if (numChoices == 0) {
		alert("You can't continue until you have entered at least one choice for this survey.");
		return false;
		} else if (numChoices == 1) {
			if (confirm("This survey only has one choice.\nAre you sure you want to continue?")) {
				return true;
			} else {
				return false;
			}
		} else {
		return true;
		}
	}
// End  -->
</script>



