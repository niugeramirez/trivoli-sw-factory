<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0


'Datos del formulario
Dim l_id
Dim l_ciudad
Dim l_codigo_postal
Dim l_idprovincia
'Dim l_balcod
'Dim l_balact
'Dim l_planro
'Dim l_balvpc
'Dim	l_balmarca
'Dim l_balconexion

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ciudades</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
if (Trim(document.datos.ciudad.value) == ""){
	alert("Debe ingresar la Ciudad.");
	document.datos.ciudad.focus();
	return;
	}
if (Trim(document.datos.codigo_postal.value) == ""){
	alert("Debe ingresar  el Codigo Postal.");
	document.datos.codigo_postal.focus();
	return;
	}	/*
else if(!stringValido(document.datos.balcod.value)){
	alert("El Código contiene caracteres inválidos.");
	document.datos.balcod.focus();
	}
else if(Trim(document.datos.baldes.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.baldes.focus();
	}
else if(!stringValido(document.datos.baldes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.baldes.focus();
	}
else if(Trim(document.datos.planro.value) == ""){
	alert("Debe ingresar una Planta.");
	document.datos.planro.focus();
	} 
else{
	var d=document.datos;
	document.valida.location = "obrassociales_06.asp?tipo=<%= l_tipo%>&id="+document.datos.id.value + "&descripcion="+document.datos.descripcion.value;
	}	*/
	valido();
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.titulo.focus();
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_id = ""
		l_ciudad = ""
		l_codigo_postal = ""
		l_idprovincia = "0"
		'l_balcod = ""
		'l_balact = ""
		'l_planro = ""
		'l_balvpc = ""
		'l_balmarca = ""		
		'l_balconexion = ""
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM ciudades "
		'l_sql = l_sql & " LEFT JOIN tkt_planta ON tkt_balanza.planro= tkt_planta.planro "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_ciudad = l_rs("ciudad")
			l_codigo_postal = l_rs("codigo_postal")
			l_idprovincia = l_rs("idprovincia")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.ciudad.focus()">
<form name="datos" action="ciudades_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="id" value="<%= l_id %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Ciudades</td>
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
					<table cellspacing="0" cellpadding="0" border="0">

					<tr>
					    <td align="right"><b>Ciudad:</b></td>
						<td>
							<input type="text" name="ciudad" size="60" maxlength="50" value="<%= l_ciudad %>">
						</td>
					</tr>		
					<tr>
					    <td align="right"><b>Codigo Postal:</b></td>
						<td>
							<input type="text" name="codigo_postal" size="60" maxlength="50" value="<%= l_codigo_postal%>">
						</td>
					</tr>	
					<tr>
					   				

						<td  align="right" nowrap><b>Provincia: </b></td>
						<td ><select name="idprovincia" size="1" style="width:200;">
								<option value=0 selected>Seleccione una Provincia</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM provincias "
								l_sql  = l_sql  & " ORDER BY provincia "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("provincia") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idprovincia.value="<%= l_idprovincia%>"</script>
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
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
