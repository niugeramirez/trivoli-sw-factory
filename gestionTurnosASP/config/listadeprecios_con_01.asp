<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_01.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_idobrasocial

l_filtro = request("filtro")
l_orden  = request("orden")
l_idobrasocial = request("idobrasocial")

if l_orden = "" then
  l_orden = " ORDER BY titulo "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Lista de Precios</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" >
<table>
    <tr>
		<th nowrap>Fecha</th>			
		<th>T&iacute;tulo</th>
		<th>Activo</th>
		<th>Acciones</th> 
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * "
l_sql = l_sql & " FROM listaprecioscabecera "
l_sql = l_sql & " WHERE idobrasocial = " & l_idobrasocial
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3" >No existen Lista de Precios cargadas.</td>
</tr>
<%else
	do until l_rs.eof
	%>	
        <tr ondblclick="Javascript:parent.abrirDialogo('dialog_lista','listadeprecios_con_02.asp?Tipo=M&cabnro=' + detalle_01_LP.cabnro.value,520,350)" onclick="Javascript:parent.Seleccionar(this,<%= l_rs("id")%>,document.detalle_01_LP.cabnro)">
			<td width="15%" align="left" nowrap><%= l_rs("fecha")%></td>
	        <td width="85%" nowrap><%= l_rs("titulo")%></td>		
			<td width="5%" nowrap><% if l_rs("flag_activo") = -1 then%>SI <% Else  %>NO<% End If %></td>
	        <td align="center" width="10%" nowrap>                    
				<a href="Javascript:parent.abrirDialogo('dialog_lista','listadeprecios_con_02.asp?Tipo=M&cabnro=' + detalle_01_LP.cabnro.value,520,350);"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>				                																				
				<a href="Javascript:parent.eliminarRegistroAJAX(document.detalle_01_LP.cabnro,'dialogAlert_lista','dialogConfirmDelete_lista');"><img src="../shared/images/Eliminar_16.png" border="0" title="Baja"></a>				
				
				<a href="Javascript:parent.abrirDialogo('dialog_cont_LP','listadeprecios_con_00.asp?id=' + document.detalle_01.cabnro.value,520,250);"><img src="../shared/images/Ecommerce-Price-Tag-icon.png" border="0" title="Lista de Precios"></a>								  				
				<a href="Javascript:parent.abrirVentana('listadepreciosdetalle_con_00.asp?id=' + detalle_01_LP.cabnro.value,'',520,200);"><img src="../shared/images/Data-List-icon_16.png" border="0" title="Detalle"></a>	
				</td>				
	    </tr>
	<%
		l_rs.MoveNext
	loop
end if
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<form name="detalle_01_LP" id="detalle_01_LP" method="post">
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
