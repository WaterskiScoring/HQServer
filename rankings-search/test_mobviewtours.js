/*
 * Javacript file for USA Waterski Tournament Listing Mobile
 * @author Mark Crone
 * @version 1.0.0
 * @date May 23, 2015
 * @copyright (c) USA Waterski
 */

function ChangeSettings()
	{
			//alert("\t ATTENTION INTERNATIONAL COMPETITOR \t \n\n THIS IS THE ALERT");
			document.getElementById('tourlisting').style.display = "none"; 		
			document.getElementById('searchsettings').style.display = "inline"; 		
			document.getElementById('Test').value = 'Mark'; 		
	}


function ReturnToListing()
	{
			document.getElementById('tourlisting').style.display = 'inline'; 		
			document.getElementById('searchsettings').style.display = 'none'; 		
	
	}


	