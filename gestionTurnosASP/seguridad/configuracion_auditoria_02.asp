<%Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
' Variables
Dim l_caudnro
Dim l_cauddes
dim	l_caudact

dim l_tipo

dim l_rs
dim l_rs1
dim l_sql

l_tipo      = Request.QueryString("tipo")
l_caudnro   = Request.QueryString("caudnro")

select Case l_tipo
	Case "A":
		 l_cauddes = ""
		 l_caudact = 0

	Case "M":
		If len(trim(l_caudnro)) = 0 then
			response.write("<script>alert('Debe seleccionar una configuracion');window.close();</script>")
		end if

		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT caudnro, "  
		l_sql = l_sql & " cauddes,  "
		l_sql = l_sql & " caudact   "
		l_sql = l_sql & " FROM  confaud"
		l_sql = l_sql & " WHERE confaud.caudnro = " & l_caudnro

		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_cauddes = l_rs("cauddes")
			l_caudact = l_rs("caudact")
		end if
		l_rs.Close
		set l_rs = nothing
end select
%>

<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n de Auditor&iacute;a</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_hora.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>
function Validar_Formulario(){
	if (document.datos.cauddes.value == "" )
		alert("Ingrese una descripcion.");
	else	
		document.datos.submit();
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" action="configuracion_auditoria_03.asp?Tipo=<%=l_tipo%>" method="post">
<input type="hidden" name="tipo" value="<%=l_tipo%>">

<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra">Config. Auditor&iacute;as</td>
		<td align="right" class="barra"><a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a></td>
	</tr>
	<tr>
		<td colspan="2" height="100%">
			<table cellpadding="0" cellspacing="0" height="100%">
				<tr>
					<td width="50%"></td>
					<td>
						<table cellpadding="0" cellspacing="0">
							<%if l_tipo = "M" then%> 
							<tr>
								<td align="right"><b>C&oacute;digo:</b></td>
								<td><input type="text" class="deshabinp" name="caudnro" size="10" maxlength="10" value="<%=l_caudnro%>" readonly></td>
							</tr>
							<%end if%>
							<tr>
								<td align="right"><b>Descripci&oacute;n:</b></td>
								<td><input type="text" name="cauddes" size="30" maxlength="30" value="<%=l_cauddes%>"></td>
							</tr>
							<tr>
								<td align="right"><b>Activo:</b></td>
								<td align="left">
									<input type="checkbox" id=checkbox1 name=caudact <% If l_caudact then %>checked<% End If %>>
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
			<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
			<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		</td>
	</tr>
</table>
</form>
</body>
</html>
