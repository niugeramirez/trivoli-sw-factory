<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: terceros_documentos_con_00.asp
'Descripción: ABM de documentos asociados a tipos de terceros
'Autor : Lisandro Moro
'Fecha: 18/02/2005

'on error goto 0

'Datos del formulario
Dim l_tipdocnro
Dim l_tipternro
Dim l_ternro

Dim l_fecvto
Dim l_oblig
Dim l_tipdocsig
Dim l_nrodoc
Dim l_tipo

'ADO
Dim l_sql
Dim l_rs

l_tipdocnro = request.querystring("tipdocnro")
l_tipternro = request.querystring("tipternro")
l_ternro = request.querystring("ternro")

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Documentos - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_fechas.js"></script> 
<script src="/ticket/shared/js/fn_mask.js"></script>
<script src="/ticket/shared/js/fn_valida.js"></script>
<script>
function Nuevo_Dialogo(w_in, pagina, ancho, alto){
	return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';help:no;status:no');
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '../shared/js/calendar.html', 16, 13);
	if (jsFecha == null){
		txt.value = '';
	}else{
		txt.value = jsFecha;
 	}
}

function Validar(){
<% select case l_tipdocnro %>
	<% case 5 %>
	if (document.datos.nrodoc.value != ""){
		if (!ValidaCuit(document.datos.nrodoc.value)){
			document.datos.nrodoc.focus();
			return;
		}
	}
<% end select %>
	if ((document.datos.fecvto.value != "")&&(!validarfecha(document.datos.fecvto))){
		 document.datos.fecvto.focus();
		 return;
	}
	document.datos.submit();
}
</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT tipdocsig, oblig, nrodoc, fecvto "
l_sql = l_sql & "  from tkt_tipodocumento "
l_sql = l_sql & "  left JOIN tkt_tipterdoc ON tkt_tipterdoc.tipdocnro = tkt_tipodocumento.tipdocnro and tkt_tipterdoc.tipternro = " & l_tipternro
l_sql = l_sql & "  LEFT JOIN tkt_terdoc ON tkt_terdoc.tipdocnro = tkt_tipodocumento.tipdocnro and tkt_terdoc.valnro = " & l_ternro & "  and tkt_terdoc.tipternro = " & l_tipternro
l_sql = l_sql & "  WHERE tkt_tipodocumento.tipdocnro = " & l_tipdocnro

rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_tipo = ""
	l_tipdocsig = l_rs("tipdocsig")
	l_fecvto = l_rs("fecvto")
	l_oblig = l_rs("oblig")
	l_nrodoc = l_rs("nrodoc")
end if

if l_nrodoc = "" or isnull(l_nrodoc) then	
	l_tipo = "A"
else
	l_tipo = ""
end if

l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="terceros_documentos_con_03.asp" method="post"  target="ifrm" >
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Documentos</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellpadding="0" cellspacing="0">
						<tr>
						    <td height="50%" align="right"><b><%= l_tipdocsig %>:</b></td>
							<td height="50%">
							<% select case l_tipdocnro %>
								<% case 5 %>
								<!--<input  type="text" name="nrodoc" size="13" maxlength="13" value="<%'= l_nrodoc %>" mask="99-99999999-9" onpaste="return tbPaste(this);" onfocus="return tbFocus(this);" onkeydown="return tbMask(this);" >-->
								<input type="text" name="nrodoc" size="13" maxlength="13" value="<%= l_nrodoc %>">
								<% case else %>
								<input type="text" name="nrodoc" size="30" maxlength="30" value="<%= l_nrodoc %>">
							<% end select %>
							</td>
						</tr>
						<tr>
						    <td height="50%" align="right" nowrap><b>Fecha Vencimiento:</b></td>
							<td height="50%" nowrap>
								<input type="text" class="habinp" name="fecvto" size="10" maxlength="10" value="<%= l_fecvto %>">
								<a href="Javascript:Ayuda_Fecha(document.datos.fecvto)">
								<img src="../shared/images/cal.gif" border="0"></a>
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
		<% call MostrarBoton ("sidebtnABM", "Javascript:Validar();","Aceptar")%>
		<a class=sidebtnABM href="Javascript:window.close();">Cancelar</a>
		
	</td>
</tr>
</table>
<iframe name="ifrm" style="visibility:hidden;" width="0" height="0"  ></iframe><!---->
<input type="hidden" name="ternro" value="<%= l_ternro %>">
<input type="hidden" name="tipternro" value="<%= l_tipternro %>">
<input type="hidden" name="tipdocnro" value="<%= l_tipdocnro %>">
<input type="hidden" name="nrodocant" value="<%= l_nrodoc %>">
<input type="hidden" name="tipo" value="<%= l_tipo %>">
<input type="hidden" name="oblig" value="<%= l_oblig %>">
</form>
<%
set l_rs = nothing
'Cn.Close
'Cn = nothing
%>
</body>
</html>
