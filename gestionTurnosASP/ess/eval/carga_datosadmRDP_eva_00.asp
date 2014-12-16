<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->

<% 
'--------------------------------------------------------------------------
'Archivo       : carga_datosadmRDP_eva_00.asp
'Descripcion   : Carga los datos de evaadmdatos (Alta y modif)
'Creacion      : 02-02-2005
'Autor         : Leticia Amadio.
'Modificacion  : 13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'				 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'--------------------------------------------------------------------------

on error goto 0
  
' Variables
' de uso local  
  Dim l_existe  
  'Dim l_gerente
  'Dim l_revisor
  Dim l_evacabnro
  Dim l_evatevnro 
  
  Dim l_horas
  Dim l_fechareunion
  Dim l_basereunion
  Dim l_evldrnro
  Dim l_tipo 
  
  
  dim l_evaevefdesde 
  dim l_evaevefhasta 
  dim l_empleado 
  dim l_emplfecalta
  dim l_empldepto 
  dim l_emplcateg 
  dim l_emplantigcateg 
  dim l_revisor  
  dim l_revcateg 
  
  dim l_htetdesde
  
  dim l_calculo
  dim l_anios 
  dim l_dias  
  dim l_meses 
  
  
  dim l_emplegactual
  dim l_empleg
' de base de datos 
  Dim l_sql
  Dim l_rs 
  Dim l_rs1
  Dim l_cm 
  
' parametros de entrada-----------------------
  l_evldrnro = Request.QueryString("evldrnro")
  l_empleg = Session("empleg")
  
Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")

'Buscar el nro de cabecera de esta evaluacion
 l_sql = "SELECT evacabnro, empleg, evadetevldor.evatevnro "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro=evadetevldor.evaluador "
 l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_evldrnro 
 rsOpen l_rs1, cn, l_sql, 0 
 if not l_rs1.EOF then 
   l_evacabnro= l_rs1("evacabnro") 
   l_emplegactual = l_rs1("empleg")
   l_evatevnro = l_rs1("evatevnro")
 end if 
 l_rs1.close 
 
' Crear registros de evadatosadm
   l_sql = "SELECT * "
   l_sql = l_sql & " FROM  evadatosadm "
   'l_sql = l_sq  & " INNER JOIN evadet ON evadet.evacabnro = .."
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evadatosadm.evldrnro" 
   l_sql = l_sql & "		AND evadetevldor.evacabnro = " & l_evacabnro
   rsOpen l_rs1, cn, l_sql, 0 
   if l_rs1.EOF then
    l_existe		= "no"
	l_horas			= "NULL" 
	l_fechareunion	= "" 
	l_basereunion	= "1"
  else 
  	l_existe		= "si"
	l_horas			= l_rs1("horas")
	l_fechareunion	= l_rs1("fechareunion")
	l_basereunion	= l_rs1("basereunion")
   end if
   l_rs1.Close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Carga de Datos Administartivos de Evaluaci&oacute;n - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>

<script>
function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha; 
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto) {
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}



function Validar_Formulario(tipo){
var errores = 0;
var deshabilitado;

		

<% 'if trim(l_empleg)<>"" and (l_empleg=l_emplegactual)  then %>
if ((errores == 0) && (document.datos.fechareunion.value == "")){
   alert("Debe ingresar una fecha.");
   document.datos.fechareunion.focus();
   errores++;   
}

if ((errores == 0) && (!validarfecha(document.datos.fechareunion))){
   document.datos.fechareunion.focus();	
   errores++;   
}
<% 'end if %>

 // document.datos.basereunion[document.datos.basereunion.selectedIndex].value

if (errores == 0){  
				//document.datos.target = "valida"
	document.datos.target = "grabar"
	document.datos.action = "grabar_datosadmRDP_eva_00.asp?Tipo="+tipo+'&evldrnro=<%=l_evldrnro%>'	
	document.datos.submit(); 
	
	if (tipo == 'M') {
		document.datos.grabado.value='M';
	} else {
		document.datos.grabado.value='G';
	}

}

}


</script>
<style>
.blanc
{
	font-size: 10;
	border-style: none;
	background : transparent;
}
.rev
{
	font-size: 10;
	border-style: none;
}
</style>
</head>

<body>

<form name="datos" action=""  method="post">
<input type="Hidden" name="terminarsecc" value="SI">
<%
	' selecciono el periodo de revision (del evento)

 l_sql = "SELECT evaevefdesde, evaevefhasta "
 l_sql = l_sql & " FROM  evaevento "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evaevenro = evaevento.evaevenro "
 l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro 
 rsOpen l_rs1, cn, l_sql, 0
 if not l_rs1.EOF then
 	l_evaevefdesde = l_rs1("evaevefdesde")
    l_evaevefhasta = l_rs1("evaevefhasta")
 end if 
 l_rs1.close 



'  buscar datos del aconsejado (nombre, categoria, depto,fecha de ingreso)

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleg, terape, ternom, terape2, ternom2, empfaltagr  " 'estrdabr, htetdesde  "
l_sql = l_sql & " ,e1.estrdabr AS depto, e2.estrdabr AS categ, categ.htetdesde "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " INNER JOIN evacab ON evacab.empleado = empleado.ternro "
l_sql = l_sql & " INNER JOIN his_estructura depto ON depto.ternro = empleado.ternro AND depto.htethasta IS NULL "
l_sql = l_sql & " INNER JOIN estructura e1 ON e1.estrnro = depto.estrnro AND e1.tenro = " & cdepartamento 
l_sql = l_sql & " INNER JOIN his_estructura categ ON categ.ternro = empleado.ternro AND categ.htethasta IS NULL "
l_sql = l_sql & " INNER JOIN estructura e2 ON e2.estrnro = categ.estrnro AND e2.tenro = " & ccategoria 
l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro
rsOpen l_rs, cn, l_sql, 0

if not l_rs.eof then
	l_empleado  = l_rs("terape") & " " &  l_rs("terape2") & ", " & l_rs("ternom") &  " "  & l_rs("ternom2")
	l_emplfecalta = l_rs("empfaltagr")
	l_empldepto = l_rs("depto") 
	l_emplcateg = l_rs("categ") 
	
	l_htetdesde = l_rs("htetdesde")
	'Calcular antig en la categoria------------
	if trim(l_htetdesde) <> "" and not isnull(l_htetdesde) then
		l_dias = DateDiff("d",l_htetdesde, date())
		l_meses = DateDiff("m",l_htetdesde, date())
		l_anios = DateDiff("yyyy",l_htetdesde, date())
		
		
		if l_dias > 364 then
			l_calculo = (l_anios * 12 ) - l_meses ' por redondeo en años???
			' if l_meses > 0 then  ' o l_calculo????????
			if l_calculo > 0 then 
				l_anios = l_anios -1
			end if
				l_dias = Int(l_dias - (l_meses * (30.5)))	' 30.416
				l_meses = Int(l_meses - (l_anios * 12 ))
			' end if	
		else
			l_anios = 0
			l_meses = l_meses - 1
			l_dias = Int(l_dias - (l_meses * (30.5)))
		end if
		
		l_emplantigcateg = ""
		if l_anios > 0 then
			if l_anios > 1 then 
				l_emplantigcateg = l_anios &" años " 
			else
				l_emplantigcateg = l_anios &" año " 
			end if
		end if 
		if l_meses > 0 then
			if l_meses > 1 then
				l_emplantigcateg = l_emplantigcateg & l_meses &" meses "
			else
				l_emplantigcateg = l_emplantigcateg & l_meses &" mes "
			end if
		end if
		'if l_dias > 0 then
			'if l_dias > 1 then
				'l_emplantigcateg = l_emplantigcateg & l_dias &" dias "
			'else
				'l_emplantigcateg = l_emplantigcateg & l_dias &" dia " 
			'end if 
		'end if
	end if
	
else
	l_empleado  = "--"
	l_emplfecalta = "--"
	l_empldepto = "--"
	l_emplcateg = "--"
	l_emplantigcateg = "--"
end if	
l_rs.Close
set l_rs=nothing
 

'  buscar datos del consejero (nombre, categoria) 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleg, terape, ternom, terape2, ternom2  " 'estrdabr, htetdesde  "
l_sql = l_sql & " ,e2.estrdabr AS categ "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evaluador = empleado.ternro "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro and  evacab.evacabnro = " & l_evacabnro 
l_sql = l_sql & " INNER JOIN his_estructura categ ON categ.ternro = empleado.ternro AND categ.htethasta IS NULL "
l_sql = l_sql & " INNER JOIN estructura e2 ON e2.estrnro = categ.estrnro AND e2.tenro = " & ccategoria 
l_sql = l_sql & " WHERE evatevnro = " & cconsejero  'cevaluador 
  ' evadetevldor.evldrnro = " & l_evldrnro & " AND.."
rsOpen l_rs, cn, l_sql, 0

if not l_rs.eof then
	l_revisor  = l_rs("terape") & " " &  l_rs("terape2") & ", " & l_rs("ternom") &  " "  & l_rs("ternom2")
	l_revcateg = l_rs("categ")
else
	l_revisor  = "--"
	l_revcateg = "--"
end if	
l_rs.Close
set l_rs=nothing
%>


<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<th colspan="3" align="left" class="th2"> Carga de Datos Administrativos RDP</th>
</tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Consejero: </b></td>
	<td> <%= l_revisor%></td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b> Categor&iacute;a de Consejero: </b></td>
 	<td><%= l_revcateg %></td> 
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b> Aconsejado: </b></td>
	<td> <%= l_empleado%></td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b> Categor&iacute;a de Aconsejado: </b></td>
 	<td><%= l_emplcateg %></td> 
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b> Linea de Servicio: </b></td>
 	<td><%= l_empldepto %></td> 
	<td>&nbsp;</td>
</tr>

<tr>
	<td align="right"><b> Antig en la categor&iacute;a: </b></td>
 	<td><%= l_emplantigcateg %></td> 
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b> Fecha de ingreso: </b></td>
 	<td><%= l_emplfecalta %></td> 
	<td>&nbsp;</td>
</tr>

<tr>
	<td align="right"><b>Per&iacute;odo de Revisi&oacute;n: </b></td>
	<td>Desde el &nbsp; <%= l_evaevefdesde%> &nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right">&nbsp;</td>
	<td>Hasta &nbsp;el &nbsp; <%= l_evaevefhasta %></td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b>Reuni&oacute;n: </b></td>
	<td> 
	<% if l_evatevnro <> cconsejero then %>
		<input class="rev" style="background : #e0e0de;" type="text" name="fechareunion" value="<%=l_fechareunion%>" size="10" readonly>&nbsp;
	<% else %> 
		<input type="text" name="fechareunion" value="<%=l_fechareunion%>" size="10" maxlength="10">&nbsp;
		<a href="Javascript:Ayuda_Fecha(document.datos.fechareunion)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a></td>
	<% end if %>		
	</td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td>&nbsp; </td> <td>&nbsp;</td>
	<% if not (l_evatevnro <> cconsejero) then %>
		<% if l_existe = "si" then      ' lo de l_tipo = "M" no me funciono!!!  %>	
			<td valign=top><a href=# onclick="Javascript:Validar_Formulario('M');">Actualizar</a>
		<%else %>
			<td valign=top><a href=# onclick="Javascript:Validar_Formulario('A');">Grabar</a>
		<%end if%>	
			<input type="text" readonly disabled name="grabado" size="1">
			</td>
	<% else %>
			<td> &nbsp;</td>
	<% end if %>

</tr>
<tr><td>
<input type="Hidden" name="horas" value=<%=l_horas%>>
<input type="Hidden" name="basereunion" value=<%=l_basereunion %>>
&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>

</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
