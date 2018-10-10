function functionLineup(theForm) {
  var totalForwards = 0;
  var totalGuards   = 0;	var totalCenters  = 0;	var totalPlayers  = 0
  var forName       = "";
  var forNameCen    = "";
	for (var i = 0; i < theForm.sCenter.length; i++) {    if (theForm.sCenter[i].checked) {      totalCenters += 1;    }  }	  if (totalCenters > 1) {    alert("1 Check Box Allowed!\n" + totalCenters + " Checked - (CENTER)");    return false;  }		for (var i = 0; i < theForm.sForward.length; i++) {
    if (theForm.sForward[i].checked) {
      totalForwards += 1;    }  }
  if (totalForwards > 2) {    alert("2 Check Boxes Allowed!\n" + totalForwards + " Checked - (FORWARD)");    return false;  }
  for (var i = 0; i < theForm.sGuard.length; i++) {    if (theForm.sGuard[i].checked) {      totalGuards += 1;    }  }
  if (totalGuards > 2) {    alert("2 Check Boxes Allowed!\n" + totalGuards + " Checked - (GUARD)");    return false;  }	/*totalPlayers = (totalCenters + totalForwards + totalGuards)	if(totalPlayers == 0) {				  alert("No Players Selected! 1 Player Required\n");    return false;			}*/	

  for (var i = 0; i < theForm.sForward.length; i++) {    forNameCen = theForm.sForward[i].value;    var arrlen = theForm.sCenter.length;    for (var j = 0; j < arrlen; j++) {     if (forNameCen == theForm.sCenter[j].value && theForm.sCenter[j].checked && theForm.sForward[i].checked) {        alert("Duplicate Player Entered - Center\n" + theForm.sForward[i].value + "");       return false;      }    }  }
  for (var i = 0; i < theForm.sForward.length; i++) {    forName = theForm.sForward[i].value;    var arrlen = theForm.sGuard.length;    for (var j = 0; j < arrlen; j++) {      if (forName == theForm.sGuard[j].value && theForm.sGuard[j].checked && theForm.sForward[i].checked) {        alert("Duplicate Player Entered - Guard\n" + theForm.sForward[i].value + "");       return false;      }    }  }	 return true;}