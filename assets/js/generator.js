var host = window.location.href.split('/generator')[0];
$("#edmType").change(function () {
	document.location.href = host + "/generator/" + $(this).val() + ".html";
});
