<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: conexion_seg_00.asp
'Descripción: 
'Autor: Lisandro Moro
'Fecha: 15/03/2005
'Modificado:
on error goto 0

Dim l_cnnro
Dim l_cndesc
Dim l_cnstring
Dim l_tipo
Dim l_sql
Dim l_rs
%>
<% 
l_tipo = request("tipo")

%>
<html>
<head>
<link href="/trivoliSwimming/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Mantenimiento Conexiones - Supervisor - RHPro &reg;</title>
</head>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script>
function Validar_Formulario(){
	if (document.datos.cndesc.value == ""){
		document.datos.cndesc.focus();
		alert("Debe ingresar la descripcion.");
		return;
	}
	if (document.datos.cnstring.value == ""){
		document.datos.cnstring.focus();
		alert("Debe ingresar el string de conexion.");
		return;	
	}
	document.datos.submit();
}
</script>
<% 
select Case l_tipo
	Case "A":
		l_cndesc = ""
		l_cnstring = ""
	Case "M":
		l_cnnro = request("cabnro")
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT cndesc, cnstring "
		l_sql  = l_sql  & "FROM conexion "
		l_sql  = l_sql  & "WHERE cnnro = " & l_cnnro
		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_cndesc = l_rs("cndesc")
			l_cnstring = l_rs("cnstring")
		end if
		l_rs.Close
		set l_rs=nothing
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="conexion_seg_03.asp?tipo=<%= l_tipo %>" method="post">
<input type="Hidden" name="cnnro" value="<%= l_cnnro %>">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
	<tr>
		<td class="th2">Datos de la Conexi&oacute;n</td>
		<td class="th2" align="right">
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
				<tr>
					<td width="50%"></td>
					<td>
						<table cellspacing="0" cellpadding="0" border="0">
							<tr>
							    <td align="right"><b>Descripci&oacute;n:</b></td>
								<td><input type="text" name="cndesc" size="35" maxlength="35" value="<%= trim(l_cndesc) %>" style="width:400;"></td>
							</tr>
							<tr>
							    <td align="right" nowrap><b>String de Conexi&oacute;n:</b></td>
								<td>
								<textarea name="cnstring" rows="5" cols="51" maxlength="255" style="width:400;"><%= trim(l_cnstring) %></textarea>
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
	    <td align="right" class="th2" colspan="2">
			<% call MostrarBoton ("sidebtnABM", "Javascript:Validar_Formulario();","Aceptar")%>
			<!--<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>-->
			<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		</td>
	</tr>
</table>
</form>
</body>
</html>
