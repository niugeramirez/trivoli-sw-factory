<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: depositos_con_02.asp
'Descripción: ABM de Depósitos
'Autor : 
'Fecha: 09/02/2005
'Modificado: 

'Datos del formulario
Dim l_depnro
Dim l_depdes
Dim l_depcod
Dim l_depmul
Dim l_deptip


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
<title><%= Session("Titulo")%>Depósitos - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
if (Trim(document.datos.depcod.value) == ""){
	alert("Debe ingresar el Código.");
	document.datos.depcod.focus();
	}
else if(!stringValido(document.datos.depcod.value)){
	alert("El Código contiene caracteres inválidos.");
	document.datos.depcod.focus();
	}
else if(Trim(document.datos.depdes.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.depdes.focus();
	}
else if(!stringValido(document.datos.depdes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.depdes.focus();
	}
else{
	var d=document.datos;
	document.valida.location = "depositos_con_06.asp?tipo=<%= l_tipo%>&depnro="+document.datos.depnro.value + "&depcod="+document.datos.depcod.value  + "&depdes="+document.datos.depdes.value;
	}	
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.depdes.focus();
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_depdes = ""
		l_depcod = ""
		l_depmul = ""
		l_deptip = ""
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_depnro = request.querystring("cabnro")
		l_sql = "SELECT depcod,depdes,depmul,deptip"
		l_sql = l_sql & " FROM tkt_deposito"
		l_sql  = l_sql  & " WHERE depnro = " & l_depnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_depdes = l_rs("depdes")
			l_depcod = l_rs("depcod")
			l_deptip = l_rs("deptip")
			l_depmul = l_rs("depmul")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.depcod.focus()">
<form name="datos" action="depositos_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="depnro" value="<%= l_depnro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Depósitos</td>
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
					    <td align="right"><b>Código:</b></td>
						<td>
							<input type="text" name="depcod" size="8" maxlength="5" value="<%= l_depcod %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Descripción:</b></td>
						<td>
							<input type="text" name="depdes" size="60" maxlength="50" value="<%= l_depdes %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Multiproducto:</b></td>
						<td>
							<input type="Checkbox" name="depmul" <% If l_depmul = -1 then  %>checked<% end if %>>
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Tipo:</b></td>
						<td>
						<select name="deptip" style="width:200px">
							<option value="C">Celda/Silo
							<option value="T">Tanque
						</select>
						<script>
							document.datos.deptip.value="<%= l_deptip %>"
						</script>
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
