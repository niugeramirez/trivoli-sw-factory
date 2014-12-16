<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!-------------------------------------------------------------------------------------------
Archivo		: filtro_campos_01.asp
Descripción : Permite realizar un filtrado 
Autor		: Lic. Fernando Favre
Fecha		: 01/2004
Modificado	: 
---------------------------------------------------------------------------------------------
-->
<% 
 Dim l_sql
 Dim l_rs
 Dim l_campos
 Dim l_tipos
 Dim l_etiqueta
 Dim l_funct
 Dim l_filtro
 Dim l_orden
 Dim l_cclave
 
 Dim l_cantidad
 Dim l_actual
 Dim l_liste(20)
 Dim l_listc(20)
 Dim l_listt(20)
 Dim l_i
 
 l_sql    	= request.Form("sql")
 l_etiqueta = request.Form("etiqueta")
 l_campos 	= request.Form("campos")
 l_tipos  	= request.Form("tipos")
 l_funct	= request.Form("funct")
 l_filtro 	= request.Form("filtro")
 l_orden	= request.Form("orden")
 l_cclave	= request.Form("campoclave")
 
 l_cantidad = 0
 do while len(l_etiqueta) > 0
 	if inStr(l_etiqueta,";") <> 0 then
    	l_actual   = left(l_etiqueta, inStr(l_etiqueta,";") - 1)
	    l_etiqueta = mid (l_etiqueta, inStr(l_etiqueta,";") + 1)
  	else
    	l_actual = l_etiqueta
		l_etiqueta = ""
	end if
  	l_cantidad = l_cantidad + 1
	l_liste(l_cantidad) = l_actual
loop
 
 l_cantidad = 0
 do while len(l_campos) > 0
 	if inStr(l_campos,";") <> 0 then
    	l_actual = left(l_campos, inStr(l_campos,";") - 1)
	    l_campos = mid (l_campos, inStr(l_campos,";") + 1)
	else
    	l_Actual = l_campos
		l_campos = ""
  	end if
	l_cantidad = l_cantidad + 1
	l_listc(l_cantidad) = l_actual
 loop
 
 l_cantidad = 0
 do while len(l_tipos) > 0
 	if inStr(l_tipos,";") <> 0 then
    	l_actual = left(l_tipos, inStr(l_tipos,";") - 1)
		l_tipos = mid (l_tipos, inStr(l_tipos,";") + 1)
	else
    	l_Actual = l_tipos
		l_tipos = ""
	end if
	l_cantidad = l_cantidad + 1
	l_listt(l_cantidad) = l_actual
 loop
 
%>
<html>	
<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Filtro - RHPro &reg;</title>
</head>
<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila, cabnro){
	if (jsSelRow != null)
  		Deseleccionar(jsSelRow);

 	document.datos.cabnro.value = cabnro;
 	fila.className = "SelectedRow";
 	jsSelRow	   = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="">
<table>
	<tr>
		<%
		l_i = 1
		do while l_i <= l_cantidad
			%>  
			<th><%= l_liste(l_i) %></th>
			<%
			l_i = l_i + 1
		loop
		%>  
	</tr>
<%
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 if l_filtro <> "" then
 	if  inStr(l_sql, "WHERE") > 0 then
		l_sql = l_sql & " AND " & l_filtro
	else
		l_sql = l_sql & " WHERE " & l_filtro
	end if
 end if
 
 l_sql = l_sql & " " & l_orden
 
 rsOpen l_rs, cn, l_sql, 0
 if l_rs.eof then
%>
	<tr>
		<td colspan="<%= l_cantidad %>">No se encontraron datos</td>
	</tr>  
<%
 else
 	do until l_rs.eof
		%>
		<tr onclick="Javascript:Seleccionar(this,<%= l_rs(l_cclave)%>)">
			<%
			l_i = 1
			do while l_i <= l_cantidad
				%>  
				<td><%= l_rs(l_listc(l_i))%></td>
				<%
				l_i = l_i + 1
			loop
			%>  
		</tr>
		<%
		l_rs.MoveNext
 	loop
 end if
%>
</table>	     
   
<%
 l_rs.close
 Set l_rs = nothing
 Cn.Close
 Set Cn = nothing
%>
<form name="datos">
	<input type="Hidden" name="cabnro" value="">
</form>
</body>
</html>
