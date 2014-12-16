<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: asignar_cascara_con_08.asp
'Descripción: Ifrm que muestra al camionero
'Autor : Raul Chinestra
'Fecha: 11/05/2005

'Datos del formulario
Dim l_ordnro
Dim l_camnro
Dim l_camcod
Dim l_camdes
Dim l_cant

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables4.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

function DatosCamionero(){
	document.valida.location = "asignar_cascara_con_07.asp?camnro=" + document.datos.camnro.value ;	
}

function actualizar_datos(pat,aco,tra){
	parent.document.datos.camcha.value = pat;
	parent.document.datos.camaco.value = aco;	
	parent.document.datos.tranro.value = tra;	
	parent.document.datos.camnro.value = document.datos.camnro.value;	
}

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_ordnro = request.querystring("ordnro")
l_camnro = request.querystring("camnro")

'response.write l_ordnro & "-"
'response.write l_camnro & "-"

if l_ordnro = 0 then
	l_camnro = 0
end if


%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">
<table cellspacing="0" cellpadding="0" border="0">
<tr>
		<td align="left">
			<select name="camnro" size="1" style="width:300;" onchange="Javascript:DatosCamionero();">
				<option value=0 selected>&laquo; Seleccione un Camionero &raquo;</option>
			<%	l_sql = "SELECT tkt_camionero.camnro, camdes, camcod "
				l_sql  = l_sql  & " FROM tkt_camionero "
				l_sql  = l_sql  & " INNER JOIN tkt_ord_cam ON tkt_ord_cam.camnro = tkt_camionero.camnro "
				l_sql  = l_sql  & " WHERE ordnro = " & l_ordnro
				'l_sql  = l_sql  & " ORDER BY entdes "
				rsOpen l_rs, cn, l_sql, 0
				
				l_cant = 0
				do until l_rs.eof 
					l_cant = l_cant + 1
				%>					
				<option value=<%= l_rs("camnro") %> > 
				<%= l_rs("camdes") %> (<%=l_rs("camcod")%>) </option>
				<%	l_rs.Movenext
				loop
				if l_cant <> 1 and l_camnro = 0 then
					l_camnro = 0
				end if
				l_rs.Close %>
			</select>
			<script> document.datos.camnro.value= "<%= l_camnro %>";
			</script>
		</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
Cn.Close
Cn = nothing
%>


</body>
</html>
