<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<%

'Archivo: perf_usr_seg_02.asp
'Descripción: Abm de Perfiles de usuario
'Autor : Alvaro Bayon
'Fecha: 21/02/2005
on error goto 0
 Dim l_sql
 Dim l_rs
 
 Dim l_perfnro
 Dim l_perfnom
 Dim l_perforden
 Dim l_perftipo
 Dim l_pol_nro
 Dim l_tipo
 
 l_tipo = request("tipo")
 l_perfnro = request("cabnro")
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Perfiles de Usuarios - Ticket</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script>
function Validar_Formulario(){
if (Trim(document.datos.perfnom.value) == ""){
	alert("Debe ingresar una Descripción.");
	document.datos.perfnom.focus();
	}
else
if(!stringValido(document.datos.perfnom.value)){
	alert("El Código contiene caracteres inválidos.");
	document.datos.perfnom.focus();
	}
else
if (document.datos.pol_nro.value == ""){
	alert("Debe seleccionar una Política de Cuenta.");
	document.datos.pol_nro.focus();
	}
else
	{
	document.datos.submit();
	}
}


</script>
<% 
select Case l_tipo
	Case "A":
		l_perfnro = ""
		l_perfnom = ""
		l_perforden = ""
		l_perftipo = 0
		l_pol_nro = ""
	Case "M":
		
		l_sql = "SELECT perfnro, perfnom,  perftipo, pol_nro "'perforden,
		l_sql  = l_sql  & "FROM perf_usr "
		l_sql  = l_sql  & "WHERE perfnro = " & l_perfnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_perfnro = l_rs("perfnro")
			l_perfnom = l_rs("perfnom")
			l_perftipo = l_rs("perftipo")
			l_pol_nro = l_rs("pol_nro")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.perfnom.focus();">
<form target="valida" name="datos" action="perf_usr_seg_03.asp?tipo=<%= l_tipo %>" method="post">
<input type="Hidden" name="perfnro" value="<%= l_perfnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
  <tr>
    <td class="th2" colspan="2" nowrap height="1">Datos del Perfil de Usuario</td>
	<td class="th2" align="right">
	 <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
  </tr>
<tr>
	<td colspan="3" height="100%">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td width="5%"></td>
				<td>
					<table cellpadding="0" cellspacing="0">
					<tr>
					    <td align="right"><b>Descripci&oacute;n:</b></td>
						<td colspan="2"><input type="text" name="perfnom" size="35" maxlength="30" value="<%= l_perfnom %>" style="width: 300px;"></td>
					</tr>
					<tr>
					    <td align="right"><b>Pol&iacute;tica de Cuenta:</b></td>
						<td colspan="2">
 						<select name="pol_nro" size="1" style="width: 300px;">
							<option value="">&laquo;Seleccione una opci&oacute;n&raquo;</option>
							<%
							l_sql = "SELECT pol_nro, pol_desc FROM pol_cuenta "

							if  Session("UserName") <> "sa" then
							  l_sql = l_sql & " WHERE pol_desc <> 'Politica Sistemas' "
  						    else
							  l_sql = l_sql & " WHERE 1 = 1 "
							end if 
  							l_sql = l_sql & " ORDER BY pol_desc "									

							response.write l_sql & "-"
							rsOpen l_rs, cn, l_sql, 0
							do until l_rs.eof
								%><option value="<%= l_rs("pol_nro") %>"><%= l_rs("pol_desc")  %></option><%
								l_rs.movenext
							loop
							l_rs.close
							%>
						</select> 
						<script>document.datos.pol_nro.value = '<%= l_pol_nro %>'</script>
						</td>
					</tr>
					</table>
				</td>
				<td width="5%"></td>
			</tr>
		</table>
	</td>
</tr>

<tr>
    <td align="right" class="th2" colspan="3" height="1">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe>
</form>
<%
	set l_rs = nothing
%>
</body>
</html>
