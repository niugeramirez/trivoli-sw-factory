<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : noved_horarias_gti_02.asp
Descripcion    : Modulo que se encarga de mostrar los datos de las nov horarias
Modificacion   :
    12/09/2003 - Scarpa D. - Coordinacion con el tablero del empleado
    07/10/2003 - Scarpa D. - Motivo no obligatorio
    20/10/2003 - Scarpa D. - correccion en la consulta SQL
    27/10/2003 - Scarpa D. - habilitar el campo descripcion
    29/10/2003 - Scarpa D. - cambio en el tamano del campo descripcion
    03/08/2005 - Scarpa D. - Solo mostrar tipo de novedades activas
	06/10/2005- Leticia A. - 
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_gnovnro
Dim l_gnovdesabr
Dim l_gnovdesext
Dim l_gnovotoa
Dim l_gtnovnro
Dim l_motnro
Dim l_gnovtipo
Dim l_gnovdesde
Dim l_gnovhasta
Dim l_gnovhoradesde
Dim l_gnovhorahasta
Dim l_gnovorden
Dim l_gnovmaxhoras

Dim l_empleado
Dim l_empleg
Dim l_ternro
dim l_eltipo

Dim l_fechadesde
Dim l_fechahasta

Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")
l_fechadesde = request.querystring("fechadesde")
l_fechahasta = request.querystring("fechahasta")
l_empleg	 = request.Querystring("empleg")

Set l_rs = Server.CreateObject("ADODB.RecordSet") 'locreo aqui para poder usarlo en las consultas de abajo

dim leg
leg = l_ess_empleg
l_ternro = l_ess_ternro

l_sql = " SELECT * FROM empleado WHERE ternro=" & l_ternro
rsOpen l_rs, cn, l_sql, 0 
l_empleado = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
l_empleg   = leg
l_rs.close


select Case l_tipo
	Case "A","TA":
		l_gnovnro = ""
		l_gnovdesabr = ""
		l_gnovdesext = ""
		l_gnovotoa = l_ess_ternro  'request.querystring("ternro")
		l_gtnovnro = ""
		l_motnro = ""
		l_gnovtipo = 1
		l_gnovdesde = ""
		l_gnovhasta = ""
		l_gnovhoradesde = ""
		l_gnovhorahasta = ""
		l_gnovorden = ""
		l_gnovmaxhoras = ""
		'l_empleado = request.querystring("empleado")
		'l_empleg = request.querystring("empleg")
	Case "M":
		l_gnovnro = request("cabnro")
		
		l_sql = "SELECT  gnovdesabr,gnovdesext, gnovotoa, gtnovnro, motnro, empleado.empleg, empleado.terape, empleado.ternom "
		l_sql = l_sql & ", gnovtipo, gnovdesde, gnovhasta, gnovhoradesde, gnovhorahasta, gnovorden, gnovmaxhoras "
		l_sql = l_sql & "FROM gti_novedad INNER JOIN empleado ON gti_novedad.gnovotoa=empleado.ternro "
		l_sql = l_sql & " WHERE gnovnro="&l_gnovnro
'		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
			l_gnovdesabr = l_rs("gnovdesabr")
			l_gnovdesext = l_rs("gnovdesext")
			l_gnovotoa = l_rs("gnovotoa")
			l_gtnovnro = l_rs("gtnovnro")
			if isNull(l_rs("motnro")) then
			   l_motnro = ""
			else
			   l_motnro = l_rs("motnro")
			end if
			l_gnovtipo = l_rs("gnovtipo")
			l_gnovdesde = l_rs("gnovdesde")
			l_gnovhasta = l_rs("gnovhasta")
			l_gnovhoradesde = l_rs("gnovhoradesde")
			l_gnovhorahasta = l_rs("gnovhorahasta")
			l_gnovorden = l_rs("gnovorden")
			l_gnovmaxhoras = l_rs("gnovmaxhoras")
			'l_empleg = l_rs("empleg")
			'l_empleado = l_rs("terape")&", "&l_rs("ternom")
		end if
		l_rs.Close
end select

%>
<html>
<head>
<link href="../<%=c_estilo%>" rel="StyleSheet" type="text/css">
<!-- <link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css"> -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Novedades Horarias - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<%
Dim l_tipAutorizacion  'Es el tipo del circuito de firmas
Dim l_HayAutorizacion  'Es para ver si las autorizaciones estan activas
Dim l_PuedeVer         'Es para ver si las autorizaciones estan activas

l_tipAutorizacion = 7  'Es del tipo novedades de gti

l_sql = "select * from cystipo "
l_sql = l_sql & "where (cystipo.cystipact = -1) and cystipo.cystipnro = " & l_tipAutorizacion 

rsOpen l_rs, cn, l_sql, 0 

l_HayAutorizacion = not l_rs.eof

l_rs.close

if l_HayAutorizacion AND (l_tipo = "M") then

  l_sql = "select cysfirautoriza, cysfirsecuencia, cysfirdestino from cysfirmas "
  l_sql = l_sql & "where cysfirmas.cystipnro = " & l_tipAutorizacion & " and cysfirmas.cysfircodext = '" & l_gnovnro & "' " 
  l_sql = l_sql & "order by cysfirsecuencia desc"

  rsOpen l_rs, cn, l_sql, 0 

  l_PuedeVer = False

  if not l_rs.eof then
    if (l_rs("cysfirautoriza") = session("UserName")) or (l_rs("cysfirdestino") = session("UserName")) then 
	   'Es una modificación del ultimo o es el nuevo que autoriza 
       l_PuedeVer = True 
    end if
  end if
  l_rs.close
  If not l_PuedeVer then
    response.write "<script>alert('No esta autorizado a ver o modificar este registro.');window.close()</script>"
	response.end
  End if
End if
%>

<script>
function horabiening(){
if (h_correcta(document.datos.gnovhoradesde1.value,document.datos.gnovhoradesde2.value)==false) 
		{ alert("Debe ingresar la hora desde o esta mal ingresada.");
		  return false;
		}
else
if (h_correcta(document.datos.gnovhorahasta1.value,document.datos.gnovhorahasta2.value)==false) 
			{alert("Debe ingresar la hora hasta, o esta mal ingresada.");
			 return false;
			}
else
	{
		if (document.datos.gnovdesde.value==document.datos.gnovhasta.value) {
			if (h_esmenor(document.datos.gnovhoradesde1.value,document.datos.gnovhoradesde2.value,document.datos.gnovhorahasta1.value,document.datos.gnovhorahasta2.value)==false)
				{ alert("La hora desde no puede ser mayor que la hora hasta.");
				  return false;
				}
			else return true;	
		}
		else {return true}
	}

}

function menorque(fecha1,fecha2){
	var f1= new Date(); 
	f1.setFullYear(fecha1.substr(6,4),fecha1.substr(3,2)-1,fecha1.substr(0,2));
	var segf1=Date.parse(f1); 

	var f2= new Date(); 
	f2.setFullYear(fecha2.substr(6,4),fecha2.substr(3,2)-1,fecha2.substr(0,2));
	var segf2=Date.parse(f2); 

	if ((segf1<segf2)||(fecha1==fecha2)){return true}
	else{return false}
}

function nuevoempleado(empleado,leg,apellido,nombre){
document.datos.gnovotoa.value= empleado;
document.datos.empleg.value= leg;
document.datos.empleado.value= apellido + ", " + nombre;
}

function Validar_Formulario()
{
<% if l_HayAutorizacion then ' Si se debe tomar autorizacion %>
// Verifico que se haya cargado la autorización 
if (((document.datos.seleccion.value == "") &&
     (document.datos.seleccion1.value == "")) &&
	 (("<%= l_tipo %>" == "A") || ("<%= l_tipo %>" == "TA")) )
    alert("Debe ingresar una autorización.");
else	
<% End If %>
if (document.datos.gnovdesabr.value == "") 
	alert("Debe ingresar una descripcíon.");
else
if (document.datos.gnovotoa.value == "") 
	alert("Debe ingresar un Empleado.");
else
if (document.datos.gtnovnro.value == "") 
	alert("Debe ingresar un tipo de Novedad.");
else
if (document.datos.gnovdesde.value == "") 
	alert("Debe ingresar la fecha desde.");
else
if (document.datos.gnovhasta.value == "") 
	alert("Debe ingresar la fecha hasta.");
else
if (validarfecha(document.datos.gnovdesde) && validarfecha(document.datos.gnovhasta)) 
{
	if (menorque(document.datos.gnovdesde.value,document.datos.gnovhasta.value)==false) 
	alert("La fecha desde no puede ser mayor que la fecha hasta.");
	else
	if (document.datos.gnovdesext.value.length > 255)
	alert('La cantidad de caracteres en el area de texto no puede superar 256');
	else
	if (document.datos.gnovtipo.value == 2) {
	if (horabiening()){
		document.datos.submit();
		}
	}
	else
	if (document.datos.gnovtipo.value == 3) 
	if (isNaN(document.datos.gnovorden.value) || (document.datos.gnovorden.value=="")) 
			alert("Debe ingresar el valor de orden, o esta mal ingresado.");
	else
	if (isNaN(document.datos.gnovmaxhoras.value) || (document.datos.gnovmaxhoras.value=="")) 
			alert("Debe ingresar el máximo a justificar, o esta mal ingresado.");
	else
		{
		document.datos.submit();
		}
	else
	{
		document.datos.submit();
	}
}
}

function habilitar(obj){
	obj.disabled = false;
	obj.className = "habinp";
}

function deshabilitar(obj){
	obj.disabled = true;
	obj.className = "deshabinp";

}

function radioclick(radio){
	if (radio==1) {
		document.datos.gnovtipo.value= "1";
		deshabilitar(document.datos.gnovorden);
		deshabilitar(document.datos.gnovmaxhoras);
		deshabilitar(document.datos.gnovhoradesde1);
		deshabilitar(document.datos.gnovhoradesde2);
		deshabilitar(document.datos.gnovhorahasta1);
		deshabilitar(document.datos.gnovhorahasta2);
	}
	if (radio==2) {
		document.datos.gnovtipo.value= "2";
		deshabilitar(document.datos.gnovorden);
		deshabilitar(document.datos.gnovmaxhoras);
		habilitar(document.datos.gnovhoradesde1);
		habilitar(document.datos.gnovhoradesde2);
		habilitar(document.datos.gnovhorahasta1);
		habilitar(document.datos.gnovhorahasta2);
	}
	if (radio==3) {
		document.datos.gnovtipo.value= "3";
		habilitar(document.datos.gnovorden);
		habilitar(document.datos.gnovmaxhoras);
		deshabilitar(document.datos.gnovhoradesde1);
		deshabilitar(document.datos.gnovhoradesde2);
		deshabilitar(document.datos.gnovhorahasta1);
		deshabilitar(document.datos.gnovhorahasta2);
	}

}
function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Firmas()  // Para llamar a control de firmas, mandandole la descripcion y demas
{
  if (document.datos.gnovdesabr.value == "")
    alert("Debe ingresar una descripción primero.")
  else	
    abrirVentana('cysfirmas_00.asp?obj=document.datos.seleccion&amp;tipo=<%= l_tipAutorizacion %>&amp;codigo=<%= l_gnovnro %>&amp;descripcion=' + document.datos.gnovdesabr.value ,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=421,height=180')
}

function actualizarDescripcion(){
  var pos = document.datos.gtnovnro.selectedIndex;
  
  document.datos.gnovdesabr.value = document.datos.gtnovnro.options[pos].innerText;
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="noved_horarias_gti_03.asp?tipo=<%= l_tipo %>&fechadesde=<%= l_fechadesde %>&fechahasta=<%= l_fechahasta %>" method="post">
<input type="Hidden" name="gnovnro" value="<%= l_gnovnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
  <tr>
    <td class="th2" colspan="2">Datos de Novedades Horarias</td>
	<td colspan="2" align="right" class="th2" valign="middle">
		&nbsp;<!-- &nbsp;&nbsp;
		<a class=sidebtnHLP href="Javascript:ayuda('<%'= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a> -->
	</td>
  </tr>
<tr>
	<td align="right">
		<br><b>Empleado:</b>
	</td>
	<td align="left" colspan="3">
		<br>
		<input type="Hidden" name="gnovotoa" value="<%= l_gnovotoa %>">	
		<input readonly type="text" name="empleg" size="10" maxlength="10" value="<%= l_empleg %>" >
		&nbsp;
		<input style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_empleado %>">
	</td>
</tr>
<tr>
    <td align="right"><b>Tipo de Novedad:</b></td>
	<td colspan="3">
	<select name="gtnovnro" size="1" onchange="javascript:actualizarDescripcion();">
	<%	l_sql = "SELECT gtnovnro, gtnovdesabr FROM gti_tiponovedad WHERE gtnovest=-1 "
'		l_rs.MaxRecords = 50
		rsOpen l_rs, cn, l_sql, 0
		do until l_rs.eof%>	
		<option value= <%= l_rs("gtnovnro") %> > 
		<%=  l_rs("gtnovdesabr") %> </option>
	<%		l_rs.Movenext
		loop
		l_rs.Close
		set l_rs= nothing 		%>	
	</select>
	<script>
		document.datos.gtnovnro.value = "<%= l_gtnovnro %>";
	</script>
	
	</td>
</tr>
<tr>
    <td align="right"><b>Motivo:</b></td>
	<td colspan="3">
	<select name="motnro" size="1">
	    <option value=""></option>
	<%	Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT motnro, motdesabr "
		l_sql  = l_sql  & "FROM gti_motivo "
'		l_rs.MaxRecords = 50
		rsOpen l_rs, cn, l_sql, 0
		do until l_rs.eof	 %>	
		<option value= <%= l_rs("motnro") %> > 
		<%= l_rs("motdesabr") %> </option>
	<%			l_rs.Movenext
		loop
		l_rs.Close
		set l_rs= nothing 		%>	
	</select>
	<script>
		document.datos.motnro.value = "<%= l_motnro %>";
	</script>	
	</td>
</tr>

<tr>
    <td align="right"><b>Desde:</b></td>
	<td>
	<input  type="text" name="gnovdesde" size="10" maxlength="10" value="<%= l_gnovdesde %>">
	<a href="Javascript:Ayuda_Fecha(document.datos.gnovdesde);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
    <td align="right"><b>Hasta:</b></td>
	<td>
	<input type="text" name="gnovhasta" size="10" maxlength="10" value="<%= l_gnovhasta %>" >
	<a href="Javascript:Ayuda_Fecha(document.datos.gnovhasta);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
</tr>

<tr>
	<td></td>
	<td colspan="3">
	<table>
	<tr>
	<td>
    <input type="Hidden" name="gnovtipo" value="<%= l_gnovtipo %>">
	<input type="radio" name="tipodia"  value="1" 
	<% if l_gnovtipo = 1 then
			response.write "Checked" 
	   end if %>  onclick="Javascript:radioclick(1);">
	Dia Completo 
	</td>														   
	<td>
	<input type="radio" name="tipodia"  value="2" 
	<% if l_gnovtipo = 2 then
			response.write "Checked" 
	   end if %>  onclick="Javascript:radioclick(2);">
	Parcial Fija 
	</td>														   
	<td>
	<input type="radio" name="tipodia"  value="3" 
	<% if l_gnovtipo = 3 then
			response.write "Checked" 
	   end if %>  onclick="Javascript:radioclick(3);">
	Parcial Variable 
	</td>														   
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="right"><b>Orden:</b></td>
	<td>
	<input type="text" name="gnovorden" size="4" maxlength="4" value="<%= l_gnovorden %>" 
	<% if l_gnovtipo <> 3 then %>
			class="deshabinp" disabled
	<% end if %>	>
	</td>
	<td align="right" class="deshab"><b>M&aacute;ximo a Just.:</b></td>
	<td>
	<input type="text" name="gnovmaxhoras" size="4" maxlength="4" value="<%= l_gnovmaxhoras %>"
	<% if l_gnovtipo <> 3 then %>
			class="deshabinp" disabled
	<% end if %>	>
	</td>
</tr>
<tr>
	<td align="right"><b>Hora Desde:</b></td>
	<td>
	<input type="text" name="gnovhoradesde1" size="2" maxlength="2" value="<%= mid(l_gnovhoradesde,1,2) %>" 
	<% if l_gnovtipo <> 2 then %>
			class="deshabinp" disabled
	<% end if %>	>
		<b>:</b>
	<input type="text" name="gnovhoradesde2" size="2" maxlength="2" value="<%= mid(l_gnovhoradesde,3,2) %>" 
	<% if l_gnovtipo <> 2 then %>
			class="deshabinp" disabled
	<% end if %>	>
	</td>
	<td align="right"><b>Hora Hasta:</b></td>
	<td>
	<input type="text" name="gnovhorahasta1" size="2" maxlength="2" value="<%= mid(l_gnovhorahasta,1,2) %>"
	<% if l_gnovtipo <> 2 then %>
			class="deshabinp" disabled
	<% end if %>	>
		<b>:</b>
	<input type="text" name="gnovhorahasta2" size="2" maxlength="2" value="<%= mid(l_gnovhorahasta,3,2) %>" 
	<% if l_gnovtipo <> 2 then %>
			class="deshabinp" disabled
	<% end if %>	>
	</td>
</tr>
<script> radioclick(<%= l_eltipo %>); </script>
<tr>
    <td align="right" height="5%"><b>Descripci&oacute;n:</b></td>
	<td colspan="3"><input type="text" name="gnovdesabr" size="30" maxlength="30" value="<%= l_gnovdesabr %>"></td>
</tr>
<tr>
    <td align="right" ><b>Desc. Extendida:</b></td>
	<td colspan="3">	    
		<textarea name="gnovdesext" rows="5" cols="40" > <%= TRIM(l_gnovdesext) %></textarea>
	</td>
</tr>
<tr>
    <td align="right" class="th2" colspan="4">
	    <input type="hidden" name="seleccion" value="">
	    <input type="hidden" name="seleccion1" value="">
<% if l_HayAutorizacion then ' Si se debe tomar autorizacion %>
		<a class=sidebtnSHW href="Javascript:Firmas()">Autorizar</a>
<% End If %>		
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>
</body>
</html>
