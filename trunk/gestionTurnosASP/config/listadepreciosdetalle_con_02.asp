<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_02.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

'Datos del formulario
Dim l_id
Dim l_titulo
Dim l_descripcion

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs


Dim l_idcab

Dim l_idpractica
Dim l_precio




l_tipo = request.querystring("tipo")
l_idcab = request.querystring("idcab")


%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Detalle de Lista de Precios</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){
/*
if (Trim(document.datos.titulo.value) == ""){
	alert("Debe ingresar el T&iacute;tulo.");
	document.datos.titulo.focus();
	return;
}


if (Trim(document.datos.descripcion.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.descripcion.focus();
	return;
}
/*
if (!stringValido(document.datos.agedes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.agedes.focus();
	return;
}

var d=document.datos;
document.valida.location = "agencias_con_06.asp?tipo=<%= l_tipo%>&agenro="+document.datos.agenro.value + "&agedes="+document.datos.agedes.value;
*/

document.datos.precio2.value = document.datos.precio.value.replace(",", ".");
if (!validanumero(document.datos.precio2, 15, 4)){
		  alert("El Precio no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.precio.focus();
		  document.datos.precio.select();
		  return;
}

valido();

}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.agedes.focus();
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_practica = ""
		l_precio = ""
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM listapreciosdetalle "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_idpractica = l_rs("idpractica")
			l_precio = l_rs("precio") 

		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.titulo.focus()">
<form name="datos" action="listadepreciosdetalle_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="id" value="<%= l_id %>">
<input type="Hidden" name="idcab" value="<%= l_idcab %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="50%">
<tr>
    <td class="th2" nowrap>Detalle de Lista de Precios</td>
</tr>
<tr>
	<td >
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">

					<tr>
						<td  align="right" nowrap><b>Practica: </b></td>
						<td ><select name="idpractica" size="1" style="width:200;">
								<option value=0 selected>Seleccione una Practica</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM practicas "
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idpractica.value="<%= l_idpractica %>"</script>
						</td>	
					</tr>						
					<tr>
					    <td align="right"><b>Precio:</b></td>
						<td>
							<input type="text" name="precio" size="40" maxlength="50" value="<%= l_precio %>">
							<input type="hidden" name="precio2" value="">	
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
    <td colspan="2" align="right" class="th">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida"  src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
