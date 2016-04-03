<meta http-equiv="refresh" content="50">
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<%
Seg = 1 / (24 * 36000000)
Plazo = seg * (cTimeOut * 60)
If (Session("Time") + Plazo < now) and (UCase(Session("Username")) <> "SUPER") then
	Session("UserName") = ""
	Session("Password") = ""
	Session("Titulo") = ""
	Session.Abandon
	'parent.parent.location = 'intro.html'
	'Response.write "<script>window.parent.parent.opener.location.reload();</script>"
	if request("menu") = "no" then
		Response.write "<script>window.parent.location = '/turnos/lanzador/lanzador2.asp';</script>"
	else
		'Response.write "<script>window.parent.parent.opener.location = '/turnos/lanzador/lanzador2.asp';</script>"
		Response.write "<script>window.parent.parent.opener.location.reload();</script>"
	end if
end if
%>	
<script>
var timerID = null;
var timerRunning = false;

function showtime (){
	timerID = setTimeout("showtime()",1000);
	timerRunning = true;
}
</script>
<body onload="javascript:showtime()" topmargin="0" leftmargin="0" bottommargin="0" rightmargin="0">
<font style="font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : xx-small;">
Cierre:<%= formatDatetime(session("Time") + Plazo,3) %><br>
Actual:<%= time %><br>
</font>
</body>
