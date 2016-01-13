<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<% 

Dim l_iduser

l_iduser = request("userid")  

%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Estructuras del Usuario - Ticket</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" action="" method="post">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="th2" align="left" colspan="2">Estructuras del Usuario</td>
	<td class="th2" colspan="2" align="right">
	&nbsp;&nbsp;&nbsp;
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0">
<tr>
	<td colspan="2" height="10">
		<br>
	</td>
</tr>
<tr>
	<td colspan="2" >
			<iframe name="ifrm" src="usuarios_seg_10.asp?iduser=<%=l_iduser%>" width="100%" height="390"></iframe> 
	</td>
</tr>
<tr>
	<td colspan="2" height="10">
		<br>
	</td>
</tr>
</table>
</form>	
</body>
</html>
