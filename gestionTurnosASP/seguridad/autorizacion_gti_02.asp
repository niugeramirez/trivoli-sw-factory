<%' Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 


Dim l_cystipnro
Dim l_cystipact 
Dim l_cystipsis
Dim l_cystipmsg
Dim l_cystipmail
Dim l_cystipnombre
Dim l_cystipprogdesc
Dim l_cystipprogdet 
Dim l_cystipprogweb 
Dim l_cystipaccion
	
Dim l_rs
Dim l_sql

Dim tipo
%>
<% 
tipo = request("tipo")
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Tipo de Autorizaci&oacute;n - Ticket</title>
</head>

<script>
function Confirmar()
{
	if (confirm('¿Esta seguro que desea Autorizar y Desautorizar con el mismo tipo de hora?'))
			return true;
	else
			return false;			
}

function Validar_Formulario()
{
if (document.datos.cystipnombre.value == "") 
	alert("Debe ingresar la nombre.");
else
	{
		document.datos.submit();
	}
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}


function disableCuatro(ori,obj,obj1,obj2,obj3)
{
	obj.disabled = !(ori.checked);
	obj1.disabled = !(ori.checked);
	obj2.disabled = !(ori.checked);
	obj3.disabled = !(ori.checked);
}

function HabilitarAut()
{
	document.datos.thautpor.disabled = true;
	document.datos.thdesautpor.disabled = true
}


function disableTres(ori,obj,obj1,obj2)
{
	obj.disabled = !(ori.checked);
	obj1.disabled = !(ori.checked);
	obj2.disabled = !(ori.checked);
}

function disableDos(ori,obj,obj1)
{
	obj.disabled = !(ori.checked);
	obj1.disabled = !(ori.checked);
}

function disableUno(ori,obj)
{
	obj.disabled = !(ori.checked);
}

function Deshabilitar()
{
//		document.datos.cystipsis.checked = false;
//	if (document.datos.l_autorizada.value = '')
//	{

//		document.datos.thdesautpor.disabled = true;
//		document.datos.combo1.disabled = true;
//		document.datos.combo2.disabled = true;
//	}
}

window.resizeTo(530,357);
</script>

<% 
select Case tipo
	Case "A":
		l_cystipnro  = ""
		l_cystipact  = ""
		l_cystipmsg  = ""
		l_cystipmail = ""
		l_cystipnombre = ""
		l_cystipprogdesc = ""
		l_cystipprogdet = ""
		l_cystipprogweb = ""
		l_cystipaccion = ""
	Case "M":
		l_cystipnro = request("cystipnro")
		
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT cystipnro, cystipnombre, "
		l_sql = l_sql & "cystipact, cystipmsg, cystipmail, "
		l_sql = l_sql & "cystipprogdesc, cystipprogdet, cystipprogweb, cystipaccion "
		l_sql = l_sql & "FROM cystipo "
		l_sql = l_sql & "WHERE cystipnro = " & l_cystipnro
		l_rs.MaxRecords = 1
		
		rsOpen l_rs, cn, l_sql, 0 
		
		if not l_rs.eof then
			l_cystipnro			= l_rs("cystipnro")
			l_cystipact			= l_rs("cystipact")
			l_cystipmsg			= l_rs("cystipmsg")
			l_cystipmail		= l_rs("cystipmail")
			l_cystipnombre		= l_rs("cystipnombre")
			l_cystipprogdesc	= l_rs("cystipprogdesc")
			l_cystipprogdet		= l_rs("cystipprogdet")
			l_cystipprogweb		= l_rs("cystipprogweb")
			l_cystipaccion      = l_rs("cystipaccion")
		end if
		
		l_rs.Close
		set l_rs = nothing
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" OnLoad="Javascript:Deshabilitar();">
<form name="datos" action="autorizacion_gti_03.asp?Tipo=<%=tipo%>" method="post" >
<input type="Hidden" name="cystipnro" value="<%= l_cystipnro %>">


<table cellspacing="1" cellpadding="0" border="0" width="100%">
<tr>
   <td class="th2" colspan="5">Datos del Tipo de Autorizaci&oacute;n</td>
</tr>
<tr>
    <td align="right"><b>C&oacute;digo:</b></td>
	<td colspan="4"><%= l_cystipnro %></td>
</tr>

<tr>
    <td align="right"><b>Nombre:</b></td>
	<td colspan="4"><input type="text" name="cystipnombre" size="30" maxlength="30" value="<%= l_cystipnombre %>"></td>
</tr>
<tr>
    <td align="right"><b>Prog. de Descripci&oacute;n:</b></td>
	<td colspan="4">
	<div style="position:relative">
	<INPUT TYPE=FILE SIZE=30 NAME="cystipprogdesc_" onchange="datos.cystipprogdesc.value = this.value;">
	<input name="cystipprogdesc" SIZE=30 type="text" value="<%= l_cystipprogdesc %>" style="position:absolute;top:0;left:0;">
	</div>
</tr>
<tr>
    <td align="right"><b>Prog. de Detalle:</b></td>
	<td colspan="4">
	<div style="position:relative">
	<INPUT TYPE=FILE SIZE=30 NAME="cystipprogdet_" onchange="datos.cystipprogdet.value = this.value;">
	<input name="cystipprogdet" SIZE=30 type="text" value="<%= l_cystipprogdet %>" style="position:absolute;top:0;left:0;">
	</div>
</tr>
<tr>
    <td align="right"><b>Prog. de Web:</b></td>
	<td colspan="4">
	<div style="position:relative">
	<INPUT TYPE=FILE SIZE=30 NAME="cystipprogweb_" onchange="datos.cystipprogweb.value = this.value;">
	<input name="cystipprogweb" SIZE=30 type="text" value="<%= l_cystipprogweb %>" style="position:absolute;top:0;left:0;">
	</div>
</tr>

<tr>
	<td align="right">
	<%if l_cystipact  then%>
	<input type="checkbox" checked id=checkbox1 name=cystipact>
	<%else%>
	<input type="checkbox" id=checkbox1 name=cystipact>
	<%end if%>
	</td>
	<td align="left"><b>Activo</b></td>
	<td align="right">
	<%if l_cystipmsg then%>
	<input type="checkbox" checked id=checkbox1 name=cystipmsg>
	<%else%>
	<input type="checkbox" id=checkbox1 name=cystipmsg>
	<%end if%>
	</td>
	<td align="left" colspan=2><b>Mensajes</b></td>
	
</tr>
<tr>
	<td align="right">
	<%if l_cystipmail  then%>
	<input type="checkbox" checked id=checkbox1 name=cystipmail>
	<%else%>
	<input type="checkbox" id=checkbox1 name=cystipmail>
	<%end if%>
	</td>
	<td align="left" colspan="4"><b>e-Mail</b></td>
</tr>


<tr>
	<td align="right">
	</td>
	<td align="left"><b>Acciones</b><br>
	<%if l_cystipaccion = "1" then %> 
	<input TYPE="radio" NAME="cystipaccion" VALUE="1" CHECKED>Sin Aviso<br>
	<%else%>
	<input TYPE="radio" NAME="cystipaccion" VALUE="1">Sin Aviso<br>
	<%end if%>
	<%if l_cystipaccion = "2" then %> 
	<input TYPE="radio" NAME="cystipaccion" VALUE="2" CHECKED>Todos Firmantes<br>
	<%else%>
	<input TYPE="radio" NAME="cystipaccion" VALUE="2" >Todos Firmantes<br>
	<%end if%>
	<%if l_cystipaccion = "3" then %> 
	<input TYPE="radio" NAME="cystipaccion" VALUE="3" CHECKED>Firmante Anterior<br>
	<%else%>
	<input TYPE="radio" NAME="cystipaccion" VALUE="3">Firmante Anterior<br>
	<%end if%>	
	<%if l_cystipaccion = "4" then %> 
	<input TYPE="radio" NAME="cystipaccion" VALUE="4" CHECKED>Primer Firmante<br>
	<%else%>
	<input TYPE="radio" NAME="cystipaccion" VALUE="4">PrimerFirmante<br>
	<%end if%>
	</td>
	<td align="left" colspan=3><b></b></td>
</tr>


</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>
</body>
</html>
