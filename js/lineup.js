function functionLineup(theForm) {
  var totalForwards = 0;
  var totalGuards   = 0;
  var forName       = "";
  var forNameCen    = "";
	for (var i = 0; i < theForm.sCenter.length; i++) {
    if (theForm.sForward[i].checked) {
      totalForwards += 1;
  if (totalForwards > 2) {
  for (var i = 0; i < theForm.sGuard.length; i++) {
  if (totalGuards > 2) {

  for (var i = 0; i < theForm.sForward.length; i++) {
  for (var i = 0; i < theForm.sForward.length; i++) {