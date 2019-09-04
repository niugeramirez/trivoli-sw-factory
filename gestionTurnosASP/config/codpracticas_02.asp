<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0


'Datos del formulario
Dim l_id
Dim l_idosocial
Dim l_idpractica
Dim l_codigo


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
<title>Codigos Practicas</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
if (Trim(document.datos.idosocial.value) == "0"){
	alert("Debe ingresar la Obra Social.");
	document.datos.idosocial.focus();
	return;
	}
if (Trim(document.datos.idpractica.value) == "0"){
	alert("Debe ingresar la Practica.");
	document.datos.idpractica.focus();
	return;
	}	
if (Trim(document.datos.codigo.value) == ""){
	alert("Debe ingresar el Codigo de Practica.");
	document.datos.codigo.focus();
	return;
	}
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
		l_id         = ""
		l_idosocial  = "0"
		l_idpractica = "0"
		l_codigo     = ""
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM codigospracticas "
		l_sql  = l_sql  & " WHERE id = " & l_id
		
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_idosocial  = l_rs("idobrasocial")
			l_idpractica = l_rs("idpractica")
			l_codigo = l_rs("codigo")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.idosocial.focus()">
<form name="datos" action="codpracticas_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="id" value="<%= l_id %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Codigos Practicas</td>
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
						<td  align="right" nowrap><b>Obra Social: </b></td>
						<td ><select name="idosocial" size="1" style="width:300;">
								<option value=0 selected>Seleccione una Obra Social</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM obrassociales "
								l_sql = l_sql & " where obrassociales.empnro = " & Session("empnro")  
								l_sql  = l_sql  & " ORDER BY descripcion "

								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idosocial.value="<%= l_idosocial%>"</script>
						</td>					
					</tr>
					<tr>
						<td  align="right" nowrap><b>Practica: </b></td>
						<td ><select name="idpractica" size="1" style="width:300;">
								<option value=0 selected>Seleccione una Practica</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM practicas "
								l_sql = l_sql & " where practicas.empnro = " & Session("empnro")  
								l_sql  = l_sql  & " ORDER BY descripcion "

								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idpractica.value="<%= l_idpractica%>"</script>
						</td>					
					</tr>
					<tr>
					    <td align="right"><b>Codigo:</b></td>
						<td>
							<input type="text" name="codigo" size="60" maxlength="50" value="<%= l_codigo%>">
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
