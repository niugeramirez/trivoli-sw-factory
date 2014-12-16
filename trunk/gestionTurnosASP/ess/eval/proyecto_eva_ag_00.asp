<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'================================================================================
'Archivo		: proyecto_eva_ag_00.asp
'Descripción	: Abm de Proyectos
'Autor			: CCRossi
'Fecha			: 30-08-2004
'Modificado		: 15-12-2004 CCRossi
' 				: 11-03-2005 - LAmadio -  Incorporarse a unproyecto existente.
'				: 23-04-2005 - LA. Cambio de revisores. 
'  				: 29-07-2005 - LA. - cambio de codigo proyecto por evento en Filtro y orden
'================================================================================
on error goto 0
' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
			' Descripci&oacute;n:; - evaproyecto.evaproynom;
  l_etiquetas = "C&oacute;digo:;Cliente:;Engagement:;Per&iacute;odo:;Fecha Desde:;Fecha Hasta:"
  l_Campos    = "evaevento.evaevenro;evacliente.evaclinom;evaengage.evaengdesabr;evaperdesabr;evaproyfdd;evaproyfht"
  l_Tipos     = "N;T;T"

' Orden
  l_Orden     = "C&oacute;digo:;Cliente:;Engagement:;Per&iacute;odo:;Fecha Desde:;Fecha Hasta:"
  l_CamposOr  = "evaevento.evaevenro;evacliente.evaclinom;evaengage.evaengdesabr;evaperdesabr;evaproyfdd;evaproyfht"

Dim l_empleg
Dim l_perfil
Dim l_ternro
Dim l_lista
Dim l_proyexistentes

Dim l_rs, l_rs1
Dim l_sql, l_sql1

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")

' l_empleg=session("empleg")
'BORRARRRR ==================================!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'l_empleg=""
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1!

if TRIM(l_empleg)="" then
	l_empleg=Request.QueryString("empleg")
end if
if trim(l_empleg)="" or isnull(l_empleg) then
	l_empleg=Request.QueryString("empleg")
end if

if trim(l_empleg)="" or isnull(l_empleg) then
	Response.Write("<script>alert('No hay usuario logeado.');window.close();</script>")
	Response.end
else
	' Buscar los datos del logeado (el ternro)
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT ternro FROM empleado WHERE empleg="& l_empleg
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then 
		l_ternro = l_rs("ternro")
	end if
	l_rs.Close
	set l_rs=nothing
	
	'Response.Write("<script>alert('"&l_ternro&"');</script>")	
	l_perfil = "empleado"
	' Buscar si es socio o gerente o que en el tipo de estructura Roles de Evaluacion
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT estrdabr "
	l_sql = l_sql & " FROM empleado "
	l_sql = l_sql & " INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro"
	l_sql = l_sql & "   AND his_estructura.tenro= " & cTenroRol
	l_sql = l_sql & "   AND his_estructura.htethasta IS NULL "
	l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro"
	l_sql = l_sql & " WHERE empleg = " & l_empleg
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_perfil = trim(l_rs("estrdabr"))
	end if
	l_rs.Close
	set l_rs=nothing
	
end if
'Response.Write("<script>alert('"&l_perfil&"');</script>")

' busco todos departamentos al que pertenece el empleado
l_lista ="0"
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT estrdabr, his_estructura.estrnro "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro"
l_sql = l_sql & "   AND his_estructura.tenro= " & cdepartamento
l_sql = l_sql & "   AND his_estructura.htethasta IS NULL "
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro"
l_sql = l_sql & " WHERE empleg = " & l_empleg
rsOpen l_rs, cn, l_sql, 0 
do while not l_rs.eof
	l_lista = l_lista & ","  & l_rs("estrnro")
l_rs.MoveNext
loop
l_rs.Close
set l_rs = nothing

' __________________________________________________________________________
' busco proyectos que tenga como linea de servicio: Depto asoc al empleado  
'    y no esten terminados 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaproyecto.evaproynro, evaproynom, evaproyecto.estrnro, evaevento.evaevenro, "
l_sql = l_sql & " evaproyfdd, evaproyfht,evaproyfin, evaclinom, evaengdesabr, evaperdesabr "
l_sql = l_sql & " FROM evaproyecto "
l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " LEFT  JOIN evaproyemp ON evaproyemp.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " LEFT JOIN evaperiodo ON evaproyecto.evapernro = evaperiodo.evapernro "
' l_sql = l_sql & " LEFT JOIN empleado ON empleado.ternro = evaproyemp.ternro "
l_sql = l_sql & " INNER JOIN evaengage  ON evaengage.evaengnro = evaproyecto.evaengnro "
l_sql = l_sql & " INNER JOIN evacliente ON evacliente.evaclinro = evaengage.evaclinro "
l_sql = l_sql & " WHERE evaproyecto.estrnro IN ("& l_lista & ")"
l_sql = l_sql & " 		AND  evaproyfin <> -1 "  ' -- ?? ver 
l_sql = l_sql & " ORDER BY evaproyecto.evaproynom, evaengage.evaengdesabr "
rsOpen l_rs, cn, l_sql, 0 

l_proyexistentes = "<select name=proyexistente style=""width:270px"">"	
l_proyexistentes = l_proyexistentes & "<option value=0>< < (Evento) Cliente-Engagement > ></option>"
do while not l_rs.eof 
	l_sql1 = " SELECT ternro FROM evaproyemp " 
	l_sql1 = l_sql1 & " WHERE evaproynro="& l_rs("evaproynro") & " AND ternro="& l_ternro ' del logueado 
	rsOpen l_rs1, cn, l_sql1, 0 
	if l_rs1.eof then 
		l_proyexistentes = l_proyexistentes & "<option value="& l_rs("evaproynro") &">" '--------
		l_proyexistentes = l_proyexistentes	&"("&l_rs("evaevenro")& ") " & l_rs("evaclinom") & " - " & l_rs("evaengdesabr") & "</option>"
	end if 
	l_rs1.Close
l_rs.MoveNext
loop 
l_rs.Close 
set l_rs = nothing 
l_proyexistentes = l_proyexistentes & "</select>  "

'l_proyexistentes = l_proyexistentes & " //<script>document.datos.proyexistente.value='<%=l_estrnro  </script>
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title>Proyectos - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function filtro(pag)
{
  abrirVentana('filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
abrirVentana('orden_param_adp_00.asp?pagina='+pag+'&lista=<%=l_orden%>&campos=<%=l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function param(){
	var chequear;
	chequear=  'ternro=<%=l_ternro%>&perfil=<%=l_perfil%>';
	return chequear;
}
   

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("proyecto_eva_ag_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value)+'&ternro=<%=l_ternro%>&perfil=<%=l_perfil%>','execl',250,150);
}

function Validar_Proyecto(){
if (document.proyecto.proyexistente.value == 0){
	alert('Debe seleccionar un proyecto');
} else {
	abrirVentanaH('proyecto_eva_ag_07.asp?ternro=<%=l_ternro%>&proyecto='+document.proyecto.proyexistente.value,"",5,5);
} 
}

</script>
</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="proyecto">
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">Proyectos</td>
          <td align="right" class="barra">
          <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('proyecto_eva_ag_02.asp?Tipo=A&ternro="&l_ternro&"&perfil="&l_perfil&"','',500,500);","Alta")%>
          <%
			if (l_perfil<>"empleado") then
           	call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'proyecto_eva_ag_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")
		   	call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('proyecto_eva_ag_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value+'&perfil="&l_perfil&"','',500,500);","Modifica")
		   	end if
          %>
          <%
			if (l_perfil<>"empleado") then
			call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('equipo_eva_00.asp?Tipo=M&evaproynro=' + document.ifrm.datos.cabnro.value+'&perfil="&l_perfil&"','',500,500);","Equipo de Trabajo")
			call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('generar_eva_00.asp?Tipo=M&evaproynro=' + document.ifrm.datos.cabnro.value+'&perfil="&l_perfil&"','',5,5);","Generar Evaluación")
			call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('cambio_evaluador_eva_00.asp?evaproynro=' + document.ifrm.datos.cabnro.value+'&perfil="&l_perfil&"','',450,400);","Cambiar Revisor")
			end if
          %>
		  &nbsp;
		  <% call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  <a class=sidebtnSHW href="Javascript:orden('proyecto_eva_ag_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('proyecto_eva_ag_01.asp')">Filtro</a>
		  <a class=sidebtnHLP href="Javascript:ayuda('<%=Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
		<tr>
		<td colspan="2"></td>
		</tr>
		<% if not (l_perfil<>"empleado") then %>
		<tr><td colspan=2>&nbsp;</td></tr>
		<tr>
			<td colspan="2" align="right">
			 <b>Proyectos existentes: </b>&nbsp;<%=l_proyexistentes%>
			 &nbsp; &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;
			 <a class=sidebtnABM href="#" onclick="Javascript:Validar_Proyecto()">&nbsp;&nbsp;&nbsp;Incorporarse al Proyecto &nbsp;&nbsp;&nbsp;</a>
			</td>
		</tr>	
		<%end if%>	
        <tr valign="top" height="100%">
          <td colspan="2" style="">
      	  <iframe name="ifrm" src="proyecto_eva_ag_01.asp?ternro=<%=l_ternro%>&perfil=<%=l_perfil%>" width="100%" height="100%" scrolling="Auto"></iframe> 
	      </td>
        </tr>

	  </table>
      </form>
</body>
</html>
