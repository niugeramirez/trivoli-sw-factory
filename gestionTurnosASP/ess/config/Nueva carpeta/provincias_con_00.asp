<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: provincias_con_00.asp
Autor: Raul Chinestra
Creacion: 22/07/2008
Descripcion: Muestra las Provincias asociada a una localidad o todas las Provincias
-----------------------------------------------------------------------------
-->
<% 

Dim l_rs
Dim l_sql
Dim l_locdes
Dim l_pronro
Dim l_disabled
Dim l_prodes

l_locdes   = request("locdes")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
'Response.write "<script>alert('"&l_locdes&"')</script>"
l_sql  =          " SELECT pronro "
l_sql  = l_sql  & " FROM int_localidad "
l_sql  = l_sql  & " WHERE locdes = '" & l_locdes & "'"
rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then 
	l_pronro = "0"
	l_disabled = ""	
else
	l_pronro = l_rs("pronro")
    l_disabled = "readonly"
end if 
'Response.write "<script>alert('"&l_pronro&"')</script>"
l_rs.close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<html>
<head>
	<title>Untitled</title>
	
<script>

function actualiza(valor){
  parent.document.datos.pro.value = valor;

}

</script>	

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" scroll=no bgcolor="#808080" >
<form name="datos">
<select <%= l_disabled %> name="pronro" size="1" style="width:240" onChange="javascript:actualiza(this.value);">
		<option value="0"></option>
<%
  	    l_sql = "SELECT pronro, prodes "
		l_sql  = l_sql  & " FROM int_provincia " 'WHERE  pronro = " & l_pronro
		rsOpen l_rs, cn, l_sql, 0 
		do until l_rs.eof 
			if CInt(l_pronro) = CInt(l_rs("pronro")) then
				l_prodes = l_rs("prodes")
			end if
		
		%>	
		<option <% if CInt(l_pronro) = CInt(l_rs("pronro")) then response.write "selected" end if%> value="<%= l_rs("prodes") %>" > 
		<%= l_rs("prodes") %></option>
		<% l_rs.Movenext
		loop
		l_rs.Close %>
	</select>
	<% if CInt(l_pronro) <> 0 then Response.write "<script>parent.document.datos.pro.value='"&l_prodes&"';</script>" end if%>
</form>

</body>
</html>
