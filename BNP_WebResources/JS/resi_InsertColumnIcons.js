//display icon and tooltio for the grid column  
function displayIconTooltip(rowData, userLCID) {
	debugger;
	var str = JSON.parse(rowData);
	var coldata = str.resi_optinemail_Value;
	var imgName = "";
	var tooltip = "";
	switch (coldata) {
		case false:
			imgName = "resi_/IMG/red.png";
			tooltip = "False";
			break;
		case true:
			imgName = "resi_/IMG/green.png";
			tooltip = "True";
			break;
		default:
			imgName = "";
			tooltip = "";
			break;
	}
	var resultarray = [imgName, tooltip];
	return resultarray;
}