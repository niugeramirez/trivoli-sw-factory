<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: sugerencias_eyp_01.asp
Descripción: Abm de postulantes
Autor : Raul Chinestra
Fecha: 22/10/2007
Modificacion: 23/11/2007 - Raul Chinestra - Se agregó una columna de Usuario
-->
<% 
on error goto 0
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_asistente
Dim l_primero
Dim l_primerob
Dim l_primeroc
Dim l_codigo
Dim l_reemplazaestrnro

l_filtro = request("filtro")
l_orden  = request("orden")
l_asistente = request("asistente")
l_codigo = request("codigo")

if l_orden = "" then
  l_orden = " ORDER BY int_sugerencia.sugfec desc "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>buques - Oleaginosa Moreno Hnos. S.A.</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
 fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,desc, codext){
    if (jsSelRow != null) {
        Deseleccionar(jsSelRow);
    };
 document.datos.cabnro.value = cabnro;
 document.datos.desc.value = codext;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
 <% 'If l_asistente = 1 then %>
    //parent.parent.ActPasos(cabnro,"Puestos",desc);
    //parent.parent.datos.pasonro.value = cabnro;
 <% 'End If %>
}

function posY(obj){
  return( obj.offsetParent==null ? obj.offsetTop : obj.offsetTop+posY(obj.offsetParent) );
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="desc" value="">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
<table border="0">
    <tr>
        <th nowrap width="100%" colspan="3" align="left">Listado de Sugerencias</th>
    </tr>
    <tr>
        <th nowrap width="10%">Fecha</th>
        <th nowrap width="10%">Usuario</th>		
        <th nowrap width="80%">Detalle</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  sugfec, sugdes, sugnro, iduser "
l_sql = l_sql & " FROM int_sugerencia"
if l_filtro <> "" then
	 l_sql = l_sql & " WHERE " & l_filtro
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Sugerencias.</td>
</tr>
<%

else
	l_primero = l_rs("sugnro")
	do while not l_rs.eof
	%>
	<tr id="<%= l_rs("sugnro") %>" onclick="Javascript:Seleccionar(this,<%= l_rs("sugnro")%>);"> 
		<td nowrap><%=l_rs("sugfec")%></td>
		<td nowrap><%=l_rs("iduser")%></td>		
		<td ><%=l_rs("sugdes")%> </td>
	</tr>
	<%
		l_rs.movenext
	loop
end if 

'...............................................................

If l_asistente = 1 then %>
    <script>    
		//alert('<%= l_primero %>');    
        //var obj = document.getElementById("<%= l_primero %>");
        //var obj = document.getElementById("<%= l_primero %>");
        //Seleccionar(obj,<%= l_primero %>);
		//alert('2');    
        //document.body.scrollTop = posY(obj); 
		//alert();
		parent.parent.ActPasos(<%= l_primero %>,"","buques");
	    parent.parent.datos.pasonro.value = <%= l_primero %>;
    </script>
<% End If 

'l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
</body>
</html>
