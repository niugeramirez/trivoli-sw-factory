<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->

<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'------------------------------------------------------------------------------------------------
'Archivo		: estadoseccion_eva_ag_00.asp
'Autor			: CCRossi
'Fecha			: 31-05-2004
'Descripcion	: estado de la seccion para el evaluador
'Modificacion	: 02-11-2004 CCRossi- verificar si la seccion es de objetivos y no se creo NINGUNO
				' no permitir terminarla
'Modificacion	: 04-11-2004 CCRossi- si Termina el evaluador (ABN) entonces copiar
				' los datos al auto y cerrar elauto tambien
'Modificacion	: 08-11-2004 CCRossi- controlar si se puede APROBAR la evaluacion
'					con las secciones OBLIGATORIAS UNICAMENTE
' 				: -03-2005 lamadio -deloitte
'-------------------------------------------------------------------------------------------------
on error goto 0
Dim l_rs
Dim l_rs1
Dim l_rs2
Dim l_sql

'locales
Dim l_evldorcargada 
Dim l_habilitado 
dim	l_evacabnro 
dim l_empleado 
dim l_cabaprobada 
dim l_readonly 
dim l_etaprogcarga
dim l_etaprogread 
dim l_depende ' guarda el nombre de la seccion de la cual depende esta y no esta TERMINADA
dim l_terminada ' habilita o no el tilde terminada
dim l_tieneobj 

dim l_evaluador	
dim l_supervisor

'parametros
dim l_evldrnro
dim l_evaseccnro
dim l_logeado ' 0: no; -1 si

l_evldrnro		= request("evldrnro")   
l_evaseccnro	= request("evaseccnro") 
l_cabaprobada	= request("cabaprobada")
l_logeado		= request("logeado")    


l_cabaprobada=""
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  evadetevldor.evacabnro, evadetevldor.evaluador, evadetevldor.evaseccnro, empleado, cabaprobada "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro"
l_sql = l_sql & " WHERE evldrnro=" & l_evldrnro
'l_sql = l_sql & " AND   evaseccnro=" & l_evaseccnro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then 
	l_evacabnro		= l_rs("evacabnro")
	l_empleado		= l_rs("empleado")
	l_evaseccnro	= l_rs("evaseccnro")
	l_evaluador		= l_rs("evaluador")
	if l_rs("cabaprobada")=-1 then
		l_cabaprobada = "checked disabled"
		l_readonly	  = "readonly disabled"
	end if	
end if
l_rs.close 
set l_rs=nothing


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  evadetevldor.evaluador "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " WHERE evadetevldor.evacabnro=" & l_evacabnro
l_sql = l_sql & " AND  (evadetevldor.evatevnro= " & cevaluador & " OR evadetevldor.evatevnro="& cconsejero &")"
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_supervisor= l_rs("evaluador")
end if
l_rs.close	
set l_rs=nothing



if l_cabaprobada = "" then
'Chequear campo readonly (a ver si hablito Cab Evaluada )
'Para cada seccion obligatoria de la cabecera... 
' ver si hay al menos evldorcargada=0 no habilito Aprobada
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  evadetevldor.evacabnro, empleado, cabaprobada,tieneobj, "
l_sql = l_sql & " evaseceta.etaprogcarga, evaseceta.etaprogread  "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro  "
l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro= evadetevldor.evaseccnro AND evasecc.evaoblig= -1 "
l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro= evadetevldor.evaseccnro AND  evaoblieva.evatevnro= evadetevldor.evatevnro "
l_sql = l_sql & " left  join evaseceta on evaseceta.evaseccnro= evadetevldor.evaseccnro "
l_sql = l_sql & "    AND  evaseceta.evatipnro= evasecc.evatipnro AND  evaseceta.evaetanro= evacab.evaetanro "
l_sql = l_sql & " WHERE evadetevldor.evacabnro = " & l_evacabnro
l_sql = l_sql & " AND   evaoblieva.evatevobli = -1 AND  evadetevldor.evldorcargada = 0  "
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
  	l_readonly = ""  
	l_cabaprobada =""
	
	' ver si hay alguna seccion con garante y hay NO acuerdo.....
	' y no está terminada, que no habilite ·Proceso Aprobado·
	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaacuerdo  "
	l_sql = l_sql & " FROM evacierre "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evacierre.evldrnro "
	l_sql = l_sql & "		 AND ( evadetevldor.evatevnro <> " & cautoevaluador
	l_sql = l_sql & "		   AND evadetevldor.evatevnro <> " & cevaluador & ")"
	l_sql = l_sql & "		   AND evadetevldor.evldorcargada <> -1" 
	l_sql = l_sql & "		   AND evadetevldor.evacabnro = " & l_evacabnro
	l_sql = l_sql & " WHERE evacierre.evaacuerdo = 0 " 
'	Response.Write "1: " & l_sql 
	rsOpen l_rs2, cn, l_sql, 0
	if not l_rs2.EOF then
		l_readonly = " readonly disabled"
	end if	
	l_rs2.close
	set l_rs2=nothing
else	
	l_readonly = " readonly disabled"
	l_etaprogcarga = l_rs("etaprogcarga")
	l_etaprogread =  l_rs("etaprogread")
	l_tieneobj=l_rs("tieneobj")
	
	if l_rs("cabaprobada")=-1 then
		l_cabaprobada = "checked disabled"
	end if	
end if
l_rs.close	
set l_rs=nothing
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>

<script>
function HabilitarProximo(){

var terminarseccion="NO"
if (parent.document.carga.document.datos.terminarsecc.value == "SI") { 
	terminarseccion="SI"
}else {
	if (parent.document.carga.document.datos.terminarsecc.value == "--") { // se fija si se cargaron resultados.
		// busco informac en el iframe oculto...
		parent.document.carga.document.terminarsecc.document.location.reload();
			// alert(parent.document.carga.document.datos.terminarsecc2.value);
		if (parent.document.carga.document.datos.terminarsecc2.value == "SI" ) {
			terminarseccion="SI"
		}
	}
}

if (terminarseccion=="SI") {
	if (confirm('¿ Da por Terminada la Sección?\n\n(Recuerde GRABAR antes de Terminar la Sección)')==true){
		if (document.all.evldorcargada.checked)	{
			document.all.evldorcargada.disabled=true;
			// actualizar edoldorcargada =-1 para el actual
			// Habilitar proximo evldrnro
			// Chequear Si estan todas las secciones obligatorias cargadas, habilitar campo aprobada
			
			abrirVentanaH('habilitar_evaluador_eva_ag_00.asp?evldrnro=<%=l_evldrnro%>&evldorcargada=-1&evaseccnro=<%=l_evaseccnro%>', '','dialogWidth:20;dialogHeight:20');
			//abrirVentana('habilitar_evaluador_eva_ag_00.asp?evldrnro=<%=l_evldrnro%>&evldorcargada=-1&evaseccnro=<%=l_evaseccnro%>', '','dialogWidth:200;dialogHeight:200');
			
			parent.actualizarcarga(<%=l_evldrnro%>,<%=l_evaseccnro%>,0,'<%=l_etaprogcarga%>','<%=l_etaprogread%>',0);	
			parent.actualizarevaluador(<%=l_evacabnro%>,<%=l_evaseccnro%>,<%=l_empleado%>);	
		}	
	}	else {
		document.all.evldorcargada.checked=false;
	}
	
} else {
		alert ('La sección no se puede terminar hasta que no se carguen todos los resultados.');
		document.all.evldorcargada.checked=false;
}
 
}


function ActualizarCab(){
	if (document.all.cabaprobada.checked) {
		document.all.cabaprobada.disabled=true;
		var r = showModalDialog('cabaprobada_eva_00.asp?evacabnro=<%=l_evacabnro%>', '','dialogWidth:20;dialogHeight:20'); 
		if (r==0){
			parent.actualizarcarga(<%=l_evldrnro%>,<%=l_evaseccnro%>,0,'<%=l_etaprogcarga%>','<%=l_etaprogread%>',0);	
			parent.actualizarevaluador(<%=l_evacabnro%>,<%=l_evaseccnro%>,<%=l_empleado%>);	
			}
		
	}	

}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" >
<table id="tabla" width="60px" height="100%">
<form name="datos" method="post">
<%	if trim(l_evldrnro)="" then 	%>
    <tr>
        <td colspan=3>Seleccione un Evaluador.</td>
    </tr>
<%
else
if trim(l_logeado)="0" then	%>
    <tr>
        <td colspan=3 align=center><b>NO habilitado.</b></td>
    </tr>
<%
else

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  * FROM evadetevldor WHERE evldrnro=" & l_evldrnro & " AND   evaseccnro=" & l_evaseccnro

rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
	
	if l_rs("habilitado")=-1 then
		l_habilitado = "CHECKED"
	else	
		l_habilitado = ""
	end if	
	if l_rs("evldorcargada")=-1 then
		l_evldorcargada = "CHECKED"
		l_terminada = "readonly disabled"
	else	
		l_evldorcargada = ""
	end if	
	
	if l_rs("habilitado")=0 then ' si no está habilitado no permito tocar terminada
		l_terminada = "readonly disabled"
	else
		l_terminada=""
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT  depende.evaseccnro, depende.titulo dependenom  "
		l_sql = l_sql & " FROM evasecc "
		l_sql = l_sql & " INNER JOIN evasecc depende ON depende.evaseccnro = evasecc.dependesecnro "
		l_sql = l_sql & " WHERE evasecc.evaseccnro = " & l_evaseccnro
		l_sql = l_sql & " AND EXISTS ("
		l_sql = l_sql & " SELECT * FROM evacab "
		l_sql = l_sql & " INNER JOIN evadetevldor det ON det.evacabnro=evacab.evacabnro "
		l_sql = l_sql & "		 AND det.evldorcargada <> -1 "
		l_sql = l_sql & "		 AND det.evacabnro  = " & l_evacabnro
		if l_tieneobj=0 then
		l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro=det.evaseccnro "
		l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro=evasecc.tipsecnro AND tipsecobj=0"
		end if
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro "
		l_sql = l_sql & "		 AND evadetevldor.evaluador  = det.evaluador"
		l_sql = l_sql & "		 AND evadetevldor.evacabnro  = " & l_evacabnro
		l_sql = l_sql & "		 AND evadetevldor.evldrnro   = " & l_evldrnro
		l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro 
		l_sql = l_sql & "		 AND det.evaseccnro = depende.evaseccnro)"
		rsOpen l_rs1, cn, l_sql, 0
		'Response.Write l_sql
		if NOT l_rs1.eof then
			l_depende = l_rs1("dependenom")
			l_terminada = " readonly disabled"
			l_evldorcargada = "" ' le saco el tilde de terminada.. no debe estar terminada!
			
			l_rs1.Close
			set l_rs1 = nothing
		else
			' verificar que si depende de una seccion, en la cual no esta el evaluador actual
			' hay que chequear si el resto la termino.
			Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT  depende.evaseccnro, depende.titulo dependenom  "
			l_sql = l_sql & " FROM evasecc "
			l_sql = l_sql & " INNER JOIN evasecc depende ON depende.evaseccnro = evasecc.dependesecnro "
			l_sql = l_sql & " WHERE evasecc.evaseccnro = " & l_evaseccnro
			l_sql = l_sql & " AND EXISTS ("
			l_sql = l_sql & " SELECT * FROM evacab "
			l_sql = l_sql & " INNER JOIN evadetevldor det ON det.evacabnro=evacab.evacabnro "
			l_sql = l_sql & "		 AND det.evldorcargada <> -1 "
			l_sql = l_sql & "		 AND det.evacabnro  = " & l_evacabnro
			l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro "
			l_sql = l_sql & "		 AND evadetevldor.evaluador  = det.evaluador"
			l_sql = l_sql & "		 AND evadetevldor.evacabnro  = " & l_evacabnro
			'l_sql = l_sql & "		 AND evadetevldor.evldrnro   = " & l_evldrnro
			l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro 
			l_sql = l_sql & "		 AND det.evaseccnro = depende.evaseccnro)"
			rsOpen l_rs1, cn, l_sql, 0
			if NOT l_rs1.eof then 
				l_depende = l_rs1("dependenom")
				l_terminada = " readonly disabled"
				l_evldorcargada = "" ' le saco el tilde de terminada.. no debe estar terminada!
				
				l_rs1.Close
				set l_rs1 = nothing
			'-----------------------------------------------------------------
			else
			
			l_rs1.Close
			set l_rs1 = nothing
			l_terminada=""
			l_depende  =""
			' verificar si la seccion es de objetivos y no se creo NINGUNO
			' no permitir terminarla
			Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evaseccnro  "
			l_sql = l_sql & " FROM evasecc "
			l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro "
			l_sql = l_sql & "		 AND evatiposecc.tipsecobj =-1 "
			l_sql = l_sql & " WHERE evasecc.evaseccnro = " & l_evaseccnro
			l_sql = l_sql & " AND EXISTS ("
			l_sql = l_sql & " SELECT * FROM evacab "
			l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro "
			l_sql = l_sql & "		 AND evadetevldor.evacabnro  = " & l_evacabnro
			l_sql = l_sql & "		 AND evadetevldor.evldrnro = " & l_evldrnro
			l_sql = l_sql & " INNER JOIN evaluaobj   ON evaluaobj.evldrnro = evadetevldor.evldrnro"
			'l_sql = l_sql & "		 AND evaluaobj.evaborrador = 0 "
			l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro=evaluaobj.evaobjnro "
			l_sql = l_sql & " WHERE evadetevldor.evaseccnro = evasecc.evaseccnro )"
			'Response.Write l_sql
			rsOpen l_rs1, cn, l_sql, 0
			if l_rs1.eof then 
				Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
				l_sql = "SELECT evaseccnro  "
				l_sql = l_sql & " FROM evasecc "
				l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro "
				l_sql = l_sql & "		 AND evatiposecc.tipsecobj =-1 "
				l_sql = l_sql & " WHERE evasecc.evaseccnro = " & l_evaseccnro
				'Response.Write l_sql
				rsOpen l_rs2, cn, l_sql, 0
				if l_rs2.EOF then
					l_terminada = ""
				else
					l_terminada = " readonly disabled"
				end if	
				l_rs2.close
				set l_rs2=nothing
			else
				l_terminada = ""	
			end if
			l_rs1.Close
			set l_rs1 = nothing
		end if
		if l_rs("evldorcargada")=-1 then
			l_terminada = "readonly disabled"
		end if	
		
		
		end if		
	end if	
%>

	<input type="Hidden" name="evldrnro" value ="<%=l_evldrnro%>">
    <tr>
		<td width="20%" align="left">&nbsp;Evaluador Habilitado:&nbsp;
			<input disabled readonly name="habilitado" type="Checkbox" <%= l_habilitado%>> 
        </td>
		<td style="bgcolor:#cccccc;">&nbsp; </td>
        <td align="right">
					&nbsp;Terminar Secci&oacute;n:&nbsp;
        	<input <%=l_terminada%> name="evldorcargada" type="Checkbox" <%= l_evldorcargada%> onclick="HabilitarProximo();">&nbsp; 
        	<%if trim(l_depende)<>"" then
        		if len(trim(l_depende))>40 then
        			l_depende=left(l_depende,40)& "..."
        		else
        			l_depende=trim(l_depende)
        		end if%>
        		(Debe terminar <%=l_depende%>)
        	<%end if%>&nbsp;
		
		<!--
		<b><%if ccodelco=-1 then%>Proceso Aprobado<%else%>Evaluaci&oacute;n Aprobada<%end if%>:</b>&nbsp;
			<input <%if l_supervisor <> l_evaluador then%>readonly disabled <%else%><%=trim(l_readonly)%> <%end if%> name="cabaprobada" type="Checkbox" <%= l_cabaprobada%> onclick="ActualizarCab();"> 
		-->
			&nbsp;
        </td>
    </tr>
<%
l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
end if ' no habilitado
end if '
cn.Close
set cn = Nothing
%>
</table>

</form>
</body>
</html>

