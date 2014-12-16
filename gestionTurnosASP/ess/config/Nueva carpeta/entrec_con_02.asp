<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: entrec_con_02.asp
'Descripción: Abm de Entregadores y recibidores
'Autor : Alvaro Bayon
'Fecha: 11/02/2005

'Datos del formulario
Dim l_entnro
Dim l_entdes
Dim l_entcod
Dim l_entact
Dim l_entrol

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
<title><%= Session("Titulo")%>Entregadores/Recibidores - Ticket - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
if (Trim(document.datos.entcod.value) == ""){
	alert("Debe ingresar el Código.");
	document.datos.entcod.focus();
	}
else if(!stringValido(document.datos.entcod.value)){
	alert("El Código contiene caracteres inválidos.");
	document.datos.entcod.focus();
	}
else if(Trim(document.datos.entdes.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.entdes.focus();
	}
else if(!stringValido(document.datos.entdes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.entdes.focus();
	}
else{
	var d=document.datos;
	document.valida.location = "entrec_con_06.asp?tipo=<%= l_tipo%>&entnro="+document.datos.entnro.value + "&entcod="+document.datos.entcod.value  + "&entdes="+document.datos.entdes.value;
	}	
}

function valido(){
	document.datos.submit();
}

function invalido(texto,foco){
	alert(texto);
	eval(foco);
	//document.datos.entdes.focus();
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto){
	return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);
	if (jsFecha == null)
		txt.value = ''
	else
		txt.value = jsFecha;
}

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
select Case l_tipo
	Case "A":
		l_entdes = ""
		l_entcod = ""
		l_entact = ""
		l_entrol = ""
	Case "M":
		l_entnro = request.querystring("cabnro")
		l_sql = "SELECT entnro,entcod,entdes,entact,entrol"
		l_sql = l_sql & " FROM tkt_entrec"
		l_sql  = l_sql  & " WHERE entnro = " & l_entnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_entcod = l_rs("entcod")
			l_entdes = l_rs("entdes")
			l_entact = l_rs("entact")
			l_entrol = l_rs("entrol")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.entcod.focus()">
<form name="datos" action="entrec_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="entnro" value="<%= l_entnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Entregadores/Recibidores</td>
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
							<input type="text" name="entcod" size="5" maxlength="3" value="<%= l_entcod %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Descripción:</b></td>
						<td>
							<input type="text" name="entdes" size="60" maxlength="50" value="<%= l_entdes %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Tipo:</b></td>
						<td>
							<input type="Radio" name="entrol" value="A" <%if l_entrol = "A" then%>checked<%end if%> >Ambos
							<input type="Radio" name="entrol" value="E" <%if l_entrol = "E" then%>checked<%end if%> >Entregador
							<input type="Radio" name="entrol" value="R" <%if l_entrol<> "E" then%>checked<%end if%> >Recibidor
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
		<% call MostrarBoton ("sidebtnABM", "Javascript:Validar_Formulario();","Aceptar")%>
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
