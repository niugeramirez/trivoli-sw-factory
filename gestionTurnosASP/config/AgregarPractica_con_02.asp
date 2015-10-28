
<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_apellido
dim l_nombre  
dim l_dni     
dim l_domicilio
dim l_idobrasocial
'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

dim l_idpractica 
dim l_idsolicitadapor
dim l_precio

Dim l_idvisita
Dim l_idpracticarealizada

l_tipo = request.querystring("tipo")
l_idvisita = request("cabnro")

l_idpracticarealizada = request("idpracticarealizada")
l_idobrasocial=request("idobrasocial")

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<!--<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Agregar Practica</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){

if (document.datos.practicaid.value == "0"){
	alert("Debe ingresar la Practica.");
	document.datos.practicaid.focus();
	return;
}

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
	document.datos.coudes.focus();
}


function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}




function calcularprecio(){


	
	document.valida.location = "agregarpractica_con_06.asp?idos=" + document.datos.idos.value + "&practicaid="+ document.datos.practicaid.value ;	
}

function actualizarprecio(p_precio){	
	document.datos.precio.value = p_precio;

}	

</script>
<% 
select Case l_tipo
	Case "A":
			l_idpractica = 0
			l_idsolicitadapor = 0
			l_precio = 0
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM practicasrealizadas "
		l_sql  = l_sql  & " WHERE id = " & l_idpracticarealizada
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_idpractica = l_rs("idpractica")
			l_idsolicitadapor = l_rs("idsolicitadapor") 
			l_precio = l_rs("precio")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.motivo.focus();">
<form name="datos" action="AgregarPractica_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="hidden" name="idvisita" value="<%= l_idvisita %>">
<input type="hidden" name="idpracticarealizada" value="<%= l_idpracticarealizada %>">
<input type="hidden" name="idos" value="<%= l_idobrasocial %>">




<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>&nbsp;</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					
											
					<tr>
						<td  align="right" nowrap><b>Practica (*): </b></td>
						<td colspan="3"><select name="practicaid" size="1" style="width:200;" onchange="calcularprecio();">
								<option value=0 selected>Seleccione una Practica</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM practicas "
								l_sql = l_sql & " where empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.practicaid.value="<%= l_idpractica %>"</script>
						</td>					
					</tr>	
					
					<tr>
						<td  align="right" nowrap><b>Solicitado por : </b></td>
						<td colspan="3"><select name="idrecursoreservable" size="1" style="width:200;">
								<option value=0 selected>Ningun Profesional</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM recursosreservables "
								l_sql = l_sql & " where empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idrecursoreservable.value="<%= l_idsolicitadapor %>"</script>							
						</td>					
					</tr>		

					<tr>
					    <td align="right"><b>Precio:</b></td>
						<td colspan="3">
							<input align="right" type="text" name="precio" size="20" maxlength="20" value="<%= l_precio %>">
							<input type="hidden" name="precio2" value="">							
						</td>
					</tr>		
						
									
					

					
					</table>
				</td>
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
