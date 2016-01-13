<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title><%= Session("Titulo")%>Principal</title>
</head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<body topmargin="0" leftmargin="0">
 <!--background="/turnos/shared/images/pi_gti.jpg"-->
 <table width="100%" height="100%" cellpadding="0" cellspacing="0">
 	<tr>
		<td width="100%" height="100%" bgcolor="#c0c0c0"></td>
	</tr>
 </table>
<div style="position:absolute; left:10; top:390">
<%
'Set rs = Server.CreateObject("ADODB.Recordset")
'sql = "select sisnom from sistema"
'rsOpen rs, cn, sql, 0
'l_sistema = rs("sisnom")
'rs.Close
'Set rs = Nothing
 %>
<iframe src="/turnos/shared/asp/timer.asp" width="90" height="30" scrolling="no" frameborder="0" style="visibility:hidden;"></iframe>
</div>
<div align="right"><b><font color="#800000" size="-2" face="Arial"><%'= l_sistema %></font></b>&nbsp;&nbsp;</div>
</body>
</html>
