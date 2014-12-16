<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: cargas_anteriores_con_02.asp
'Descripción: ABM Cargas Anteriores
'Autor : Gustavo Manfrin
'Fecha: 07/08/2005
'Modificado: 

'Datos del formulario

on error goto 0 

Dim l_carconnro
Dim l_lugnro
Dim l_pronro
Dim l_prodes
Dim l_procod

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Cargas Anteriores - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){

if (document.datos.lugnro.value== 0){

	alert("Debe ingresar el Lugar.");
	document.datos.lugnro.focus();
	return;
}

if (document.datos.pronro.value == 0){
	alert("Debe ingresar el Producto.");
	document.datos.procod.focus();
	return;
}

var d=document.datos;
	
document.valida.location = "cargas_anteriores_con_06.asp?tipo=<%= l_tipo%>&carconnro="+document.datos.carconnro.value + "&pronro="+document.datos.pronro.value+ "&lugnro="+document.datos.lugnro.value;

}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.lugnro.focus();
}

</script>
<% 
select Case l_tipo
	Case "A":
       l_lugnro = ""
       l_pronro = ""
       l_prodes = ""
       l_procod = ""
       Set l_rs = Server.CreateObject("ADODB.RecordSet")
	   
 	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_carconnro = request.querystring("cabnro")
        l_sql = "SELECT tkt_cargasconf.lugdesnro, tkt_cargasconf.pronro, "
        l_sql = l_sql & " tkt_lugar.lugcod, tkt_producto.prodes, tkt_producto.procod "
        l_sql = l_sql & " FROM tkt_cargasconf "
        l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_cargasconf.lugdesnro= tkt_lugar.lugnro "
        l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_cargasconf.pronro= tkt_producto.pronro "
        l_sql = l_sql & " WHERE tkt_cargasconf.carconnro = " & l_carconnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
           l_lugnro = l_rs("lugdesnro")
           l_pronro = l_rs("pronro")
           l_prodes = l_rs("prodes")
           l_procod = l_rs("procod")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.pronro.focus()">
<form name="datos" action="cargas_anteriores_con_03.asp?tipo=<%= l_tipo %>&lugnro=<%= l_lugnro %>" method="post" target="valida">
<input type="Hidden" name="carconnro" value="<%= l_carconnro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Cargas Anteriores</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" >
					<tr>
						<td align="right" nowrap><b>Lugar Destino:</b></td>
						<td>
							<select name="lugnro" size="1" style="width:200;"  >
							<option value="" selected>&laquo; Seleccione un Lugar &raquo;</option>
							<%	l_sql = "SELECT lugnro, lugdes, lugcod "
								l_sql  = l_sql  & " FROM tkt_lugar "
								l_sql  = l_sql  & " ORDER BY lugdes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof %>	
									<option value=<%= l_rs("lugnro") %> > 
									<%= l_rs("lugdes") %> (<%=l_rs("lugcod")%>) </option>
 									<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script> document.datos.lugnro.value= "<%= l_lugnro %>";</script>
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Producto Cargado:</b></td>						
						<td>
							<select name="pronro" size="1" style="width:400;" <%'= l_claseCombo %>>
							<option value="" selected>&laquo; Seleccione producto &raquo;</option>
							<%	l_sql = "SELECT pronro, prodes, procod "
								l_sql  = l_sql  & " FROM tkt_producto "
'								l_sql  = l_sql  & " WHERE proest <> 0 "
								l_sql  = l_sql  & " ORDER BY prodes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof %>	
									<option value=<%= l_rs("pronro") %> > 
									<%= l_rs("prodes") %> (<%=l_rs("procod")%>) </option>
 									<%	l_rs.Movenext
								loop
								'l_rs.Close %>
							</select>
							<script> document.datos.pronro.value= "<%= l_pronro %>"</script>
						</td>
						</tr>
					</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
