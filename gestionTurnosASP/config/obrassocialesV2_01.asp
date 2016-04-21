<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 


Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY id"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="../shared/js/fn_windows.js"></script>
<script src="../shared/js/fn_confirm.js"></script>
<script src="../shared/js/fn_ayuda.js"></script>

<head>
<link href="../ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Obras Sociales</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        
        <th>Descripci&oacute;n</th>
		<th>Acciones</th> 
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * "
l_sql = l_sql & " FROM obrassociales "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
  l_sql = l_sql & " and obrassociales.empnro = " & Session("empnro")   
else
  l_sql = l_sql & " where obrassociales.empnro = " & Session("empnro")   
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen Obras Sociales</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirDialogo('dialog','obrassocialesV2_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,600,300)" onclick="Javascript:parent.Seleccionar(this,<%= l_rs("id")%>,document.detalle_01.cabnro)">
            <td width="20%" nowrap><%= l_rs("descripcion")%></td>
	        <td align="center" width="10%" nowrap>                    
				<a href="Javascript:parent.abrirDialogo('dialog','obrassocialesV2_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,600,300);"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>				                								
				<a href="Javascript:parent.eliminarRegistroAJAX(document.detalle_01.cabnro,'dialogAlert','dialogConfirmDelete');"><img src="../shared/images/Eliminar_16.png" border="0" title="Baja"></a>				
				<a href="Javascript:parent.abrirDialogo('dialog_cont_LP','listadeprecios_con_00.asp?id=' + document.detalle_01.cabnro.value,520,250);"><img src="../shared/images/Ecommerce-Price-Tag-icon.png" border="0" title="Lista de Precios"></a>								  				
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
<form name="detalle_01" id="detalle_01" method="post">
	<input type="hidden" id="cabnro" name="cabnro" value="0">
	<input type="hidden" name="orden" value="<%= l_orden %>">
	<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
