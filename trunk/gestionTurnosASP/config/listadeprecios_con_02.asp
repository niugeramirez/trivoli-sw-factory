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
Dim l_fecha

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

Dim l_flag_activo
Dim l_idobrasocial
Dim l_flag_percoplistprec

l_tipo = request.querystring("tipo")
l_idobrasocial = request.querystring("idobrasocial")

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Lista de Precios</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>

<!-- Comienzo Datepicker -->
<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css">
<script src="../js/jquery-1.8.0.js"></script>
<script src="../js/jquery-ui.js"></script>  
<script src="../js/jquery.ui.datepicker-es.js"></script>
<script>
$(function () {
$.datepicker.setDefaults($.datepicker.regional["es"]);
$("#datepicker").datepicker({
firstDay: 1
});

		
$( "#fecha" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->

<script>
function Validar_Formulario(){

if ((document.datos.fecha.value == "")&&(!validarfecha(document.datos.fecha))){
	 document.datos.fecha.focus();
	 return;
}


if (Trim(document.datos.titulo.value) == ""){
	alert("Debe ingresar el Titulo.");
	document.datos.titulo.focus();
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

function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function enable_Copy_List_Prec() {
	if (document.datos.flag_percoplistprec.value == "0") {
			//document.datos.idobrasocial.readOnly = false;
			//document.datos.idobrasocial.className = 'habinp';			
			document.datos.lpcab.disabled = false;							
		}
		else {
			//document.datos.idobrasocial.readOnly = true;
			//document.datos.idobrasocial.className = 'deshabinp';		
			document.datos.lpcab.disabled = true;	
		}	
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_titulo = ""
		l_fecha = ""
		l_flag_activo = "0"
		l_flag_percoplistprec = "0"
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " ,( select count(*) 	 from listapreciosdetalle  where idlistaprecioscabecera = listaprecioscabecera.id) as flag_tiene_detalle"
		l_sql = l_sql & " FROM listaprecioscabecera "
		l_sql  = l_sql  & " WHERE id = " & l_id
		
		'response.write l_sql & "<br>"
		
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_titulo = l_rs("titulo")
			l_fecha = l_rs("fecha") 
			l_flag_activo = l_rs("flag_activo")
			l_flag_percoplistprec = l_rs("flag_tiene_detalle")
		end if
		l_rs.Close
end select 
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.fecha.focus()">
<form name="datos" action="listadeprecios_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="id" value="<%= l_id %>">
<input type="Hidden" name="idobrasocial" value="<%= l_idobrasocial %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Lista de Precios</td>
	<td class="th2" align="right">

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
					    <td align="right" nowrap width="0"><b>Fecha:</b></td>
						<td align="left" nowrap width="0" >
						    <input type="text" id="fecha" name="fecha" size="10" maxlength="10" value="<%= l_fecha %>">
							
						</td>																	
					</tr>							
					<tr>
					    <td align="right"><b>T&iacute;tulo:</b></td>
						<td>
							<input type="text" name="titulo" size="50" maxlength="100" value="<%= l_titulo %>">
						</td>
					</tr>		
					<tr>
						<td  align="right" nowrap><b>Activo: </b></td>
						<td ><select name="activo" size="1" style="width:250;">
								<option value=0 selected>No</option>
								<option value=-1 selected>Si</option>								
							</select>
							<script>document.datos.activo.value="<%= l_flag_activo %>"</script>
						</td>		
					</tr>		
					<tr>
						<td  align="right" nowrap><b>Copiar precios de: </b></td>
						<td >
							<input type="Hidden" name="flag_percoplistprec" value="<%= l_flag_percoplistprec %>">							
							<select name="lpcab" size="1" style="width:250;">
								<option value=0 selected>Seleccione una Lista de Precios</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT * "
								l_sql = l_sql & " FROM listaprecioscabecera "								
								l_sql  = l_sql  & " ORDER BY titulo "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("titulo") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>enable_Copy_List_Prec();</script>
							
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
