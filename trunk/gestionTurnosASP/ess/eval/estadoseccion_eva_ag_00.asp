<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
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
'				: 10-07-2006 - LA. - Cuando mira si las seccion dependiente termino, fijarse solo en roles obligatorios y fijarse que la sección dependiente sea tb Obligatoria.
'				: 10-07-2006 - LA. - Se saco la resticcion de que el que podia aprobar la Evaluacion sea el Supervisor
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
l_sql = "SELECT  evadetevldor.evacabnro, evadetevldor.evaluador, evadetevldor.evaseccnro, empleado, cabaprobada FROM evadetevldor INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro WHERE evldrnro=" & l_evldrnro
'l_sql = l_sql & " AND   evaseccnro=" & l_evaseccnro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then
	l_evacabnro		= l_rs("evacabnro")
	l_empleado		= l_rs("empleado")
	l_evaseccnro	= l_rs("evaseccnro")
	l_evaluador		= l_rs("evaluador")
	if cint(l_rs("cabaprobada"))=-1 then
		l_cabaprobada = " checked disabled"
		l_readonly    = " readonly disabled"
	end if	
end if
l_rs.close	
set l_rs=nothing


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  evadetevldor.evaluador FROM evadetevldor WHERE evadetevldor.evacabnro=" & l_evacabnro & " AND   evadetevldor.evatevnro= " & cevaluador
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_supervisor= l_rs("evaluador")
end if
l_rs.close	
set l_rs=nothing

if l_cabaprobada = "" then
'Chequear campo readonly (a ver si hablito Cab Evaluada )
'Para cada seccion obligatoria de la cabecera... 
'ver si hay al menos evldorcargada=0 no habilito Aprobada

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  evadetevldor.evacabnro, evadetevldor.evatevnro, empleado, cabaprobada,tieneobj, evaseceta.etaprogcarga, evaseceta.etaprogread FROM evadetevldor INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro "
l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro= evadetevldor.evaseccnro AND evasecc.evaoblig= -1 INNER JOIN evaoblieva ON evaoblieva.evaseccnro= evadetevldor.evaseccnro AND  evaoblieva.evatevnro= evadetevldor.evatevnro"
l_sql = l_sql & " left  join evaseceta on evaseceta.evaseccnro= evadetevldor.evaseccnro AND  evaseceta.evatipnro= evasecc.evatipnro AND  evaseceta.evaetanro= evacab.evaetanro WHERE evadetevldor.evacabnro = " & l_evacabnro & " AND   evaoblieva.evatevobli = -1 AND evadetevldor.evldorcargada = 0" 
rsOpen l_rs, cn, l_sql, 0
'response.write l_sql & "<br>"
if l_rs.eof then
  	l_readonly    =""
	l_cabaprobada =""
	
	' ver si hay alguna seccion con garante y hay NO acuerdo.....
	' y no está terminada, que no habilite Proceso Aprobado

	' si hubo desacuerdo por parte del supervisado
	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaacuerdo FROM evacierre INNER JOIN evadetevldor ON evadetevldor.evldrnro=evacierre.evldrnro "
	l_sql = l_sql & " AND evadetevldor.evatevnro = " & cautoevaluador 
	l_sql = l_sql & " AND evadetevldor.evacabnro = " & l_evacabnro 
	l_sql = l_sql & " WHERE evacierre.evaacuerdo = 0 " 
	rsOpen l_rs2, cn, l_sql, 0
	if not l_rs2.EOF then
		l_rs2.close
		
		' si el garante no la aprobo que no habilite
		l_sql = "SELECT evaacuerdo FROM evacierre INNER JOIN evadetevldor ON evadetevldor.evldrnro=evacierre.evldrnro "
		l_sql = l_sql & " AND evadetevldor.evatevnro = " & cgarante 
		l_sql = l_sql & " AND evadetevldor.evacabnro = " & l_evacabnro 
		l_sql = l_sql & " WHERE evadetevldor.evldorcargada = 0 " 
		rsOpen l_rs2, cn, l_sql, 0
		if not l_rs2.EOF then
			l_readonly = " readonly disabled"
		end if
		l_rs2.close
	else
		l_rs2.close	
	end if
	set l_rs2=nothing
else	

	l_readonly = " readonly disabled"		

	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT COUNT(evldrnro) as noterminados FROM evadetevldor INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro "
	l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro= evadetevldor.evaseccnro AND evasecc.evaoblig= -1 INNER JOIN evaoblieva ON evaoblieva.evaseccnro= evadetevldor.evaseccnro AND  evaoblieva.evatevnro= evadetevldor.evatevnro"
	l_sql = l_sql & " left  join evaseceta on evaseceta.evaseccnro= evadetevldor.evaseccnro AND  evaseceta.evatipnro= evasecc.evatipnro AND  evaseceta.evaetanro= evacab.evaetanro WHERE evadetevldor.evacabnro = " & l_evacabnro & " AND   evaoblieva.evatevobli = -1 AND evadetevldor.evldorcargada = 0" 
 	rsOpen l_rs2, cn, l_sql, 0
	if not l_rs2.EOF then
		if cint(l_rs2("noterminados")) = 1 and cint(l_rs("evatevnro")) = cgarante then
			l_rs2.close
			' si solo hay una seccion obligatoria or terminar y es del garante
			' hay que ver si hay acuerdo.
			l_sql = "SELECT evaacuerdo FROM evacierre INNER JOIN evadetevldor ON evadetevldor.evldrnro=evacierre.evldrnro "
			l_sql = l_sql & " AND evadetevldor.evatevnro = " & cautoevaluador 
			l_sql = l_sql & " AND evadetevldor.evacabnro = " & l_evacabnro 
			l_sql = l_sql & " WHERE evacierre.evaacuerdo = -1 " 
			rsOpen l_rs2, cn, l_sql, 0
			if not l_rs2.EOF then
				l_readonly=""
				l_rs2.close
			else
				l_rs2.close 
				l_readonly = " readonly disabled"				
			end if
		end if
	end if
	set l_rs2=nothing
	
	l_etaprogcarga = l_rs("etaprogcarga")
	l_etaprogread =  l_rs("etaprogread")
	l_tieneobj    =  cint(l_rs("tieneobj"))
	 	 
	 
	if cint(l_rs("cabaprobada"))=-1 then
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
	if (confirm('¿ Da por Terminada la Sección?\n\n(Recuerde GRABAR antes de Terminar la Sección)')==true)
 	{
	if (document.all.evldorcargada.checked)
	{
		document.all.evldorcargada.disabled=true;
		// actualizar edoldorcargada =-1 para el actual
		// Habilitar proximo evldrnro
		// Chequear Si estan todas las secciones obligatorias cargadas, habilitar campo aprobada
// var r = showModalDialog('habilitar_evaluador_eva_ag_00.asp?evldrnro=<%=l_evldrnro%>&evldorcargada=-1&evaseccnro=<%=l_evaseccnro%>', '','dialogWidth:200;dialogHeight:200'); 
		abrirVentanaH('habilitar_evaluador_eva_ag_00.asp?evldrnro=<%=l_evldrnro%>&evldorcargada=-1&evaseccnro=<%=l_evaseccnro%>' ,'',250,250);
		var r =0;
		if (r==0){
			parent.actualizarcarga(<%=l_evldrnro%>,<%=l_evaseccnro%>,0,'<%=l_etaprogcarga%>','<%=l_etaprogread%>',0);	
			parent.actualizarevaluador(<%=l_evacabnro%>,<%=l_evaseccnro%>,<%=l_empleado%>);	
			}
		
	}
	}
	else
		document.all.evldorcargada.checked=false;
}

function ActualizarCab(){
	if (document.all.cabaprobada.checked)
	{
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
<table id="tabla">
<form name="datos" method="post">
<%
	if trim(l_evldrnro)="" then
	%>
    <tr>
        <td colspan=8>Seleccione un Evaluador.</td>
    </tr>
<%
else
if trim(l_logeado)="0" then
	%>
    <tr>
        <td colspan=8 align=center><b>NO habilitado.</b></td>
    </tr>
<%
else

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  * FROM evadetevldor WHERE evldrnro=" & l_evldrnro & " AND   evaseccnro=" & l_evaseccnro

rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
	
	if cint(l_rs("habilitado"))=-1 then
		l_habilitado = "CHECKED"
	else	
		l_habilitado = ""
	end if	
	if cint(l_rs("evldorcargada"))=-1 then
		l_evldorcargada = "CHECKED"
		l_terminada = " readonly disabled"
	else	
		l_evldorcargada = ""
	end if	
	
	if cint(l_rs("habilitado"))=0 then ' si no está habilitado no permito tocar terminada
		l_terminada = " readonly disabled"
	else
		l_terminada=""
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT  depende.evaseccnro, depende.titulo dependenom FROM evasecc INNER JOIN evasecc depende ON depende.evaseccnro = evasecc.dependesecnro "
		l_sql = l_sql & " WHERE evasecc.evaseccnro = " & l_evaseccnro & " AND depende.evaoblig= -1 " ' fijarse que la seccion de la cual depende sea obligatoria
		l_sql = l_sql & " AND EXISTS ("
		l_sql = l_sql & " SELECT * FROM evacab INNER JOIN evadetevldor det ON det.evacabnro=evacab.evacabnro "
		l_sql = l_sql & "   AND det.evldorcargada <> -1 AND det.evacabnro  = " & l_evacabnro
				' fijarse solo en roles obligatorios 
		l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro=det.evaseccnro "
		l_sql = l_sql & "   AND evaoblieva.evatevnro=det.evatevnro AND evaoblieva.evatevobli =-1 "
		if cint(l_tieneobj)=0 then
		l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro=det.evaseccnro INNER JOIN evatiposecc ON evatiposecc.tipsecnro=evasecc.tipsecnro AND tipsecobj=0 "
		end if
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro AND evadetevldor.evaluador  = det.evaluador AND evadetevldor.evacabnro  = " & l_evacabnro & " AND evadetevldor.evldrnro   = " & l_evldrnro
		l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro & " AND det.evaseccnro = depende.evaseccnro)"
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
			l_sql = "SELECT  depende.evaseccnro, depende.titulo dependenom FROM evasecc INNER JOIN evasecc depende ON depende.evaseccnro = evasecc.dependesecnro "
			l_sql = l_sql & " WHERE evasecc.evaseccnro = " & l_evaseccnro & " AND depende.evaoblig= -1 " ' fijarse que la seccion de la cual depende sea obligatoria
			l_sql = l_sql & " AND EXISTS ( "
			l_sql = l_sql & " SELECT * FROM evacab INNER JOIN evadetevldor det ON det.evacabnro=evacab.evacabnro "
			l_sql = l_sql & " AND det.evldorcargada <> -1 AND det.evacabnro=" & l_evacabnro
					' fijarse solo en roles obligatorios 
			l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro=det.evaseccnro "
			l_sql = l_sql & "     AND evaoblieva.evatevnro=det.evatevnro AND evaoblieva.evatevobli =-1 "
			l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro AND evadetevldor.evaluador  = det.evaluador AND evadetevldor.evacabnro  = " & l_evacabnro
			l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro  & "	AND det.evaseccnro = depende.evaseccnro)"
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
			l_sql = "SELECT evaseccnro FROM evasecc INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro AND evatiposecc.tipsecobj =-1 "
			l_sql = l_sql & " WHERE evasecc.evaseccnro = " & l_evaseccnro & " AND EXISTS ("
			l_sql = l_sql & " SELECT * FROM evacab INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro AND evadetevldor.evacabnro  = " & l_evacabnro & "	AND evadetevldor.evldrnro = " & l_evldrnro
			l_sql = l_sql & " INNER JOIN evaluaobj   ON evaluaobj.evldrnro = evadetevldor.evldrnro INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro=evaluaobj.evaobjnro WHERE evadetevldor.evaseccnro = evasecc.evaseccnro )"
			'Response.Write l_sql
			rsOpen l_rs1, cn, l_sql, 0
			if l_rs1.eof then
				Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
				l_sql = "SELECT evaseccnro FROM evasecc INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro AND evatiposecc.tipsecobj =-1 WHERE evasecc.evaseccnro = " & l_evaseccnro
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
		if cint(l_rs("evldorcargada"))=-1 then
			l_terminada = " readonly disabled"
		end if	
		
		
		end if		
	end if	

	%>
	<input type="Hidden" name="evldrnro" value ="<%=l_evldrnro%>">
    <tr>
		<td width="70%" align="center">Evaluador Habilitado:&nbsp;
			<input disabled readonly name="habilitado" type="Checkbox" <%= l_habilitado%> > 
			&nbsp;Secci&oacute;n Terminada:&nbsp;
        	<input <%=l_terminada%> name="evldorcargada" type="Checkbox" <%= l_evldorcargada%> onClick="HabilitarProximo();"> 
        	<%if trim(l_depende)<>"" then
        		if len(trim(l_depende))>40 then
        			l_depende=left(l_depende,40)& "..."
        		else
        			l_depende=trim(l_depende)
        		end if%>
        		(Debe terminar <%=l_depende%>)
        	<%end if%>
        </td>
        <td align="right"><b><%if ccodelco=-1 then%>Proceso Aprobado<%else%>Evaluaci&oacute;n Aprobada<%end if%>:</b>&nbsp;

<input <%= l_readonly %> name="cabaprobada" type="Checkbox" <%= l_cabaprobada%> onClick="ActualizarCab();"> 
<!-- SE comenta- esto es para Codelco input <%'if cdbl(l_supervisor)<>cdbl(l_rs("evaluador")) then %>readonly disabled <%'else%><%'=trim(l_readonly)%> <%'end if%> name="cabaprobada" type="Checkbox" <%'= l_cabaprobada%> onclick="ActualizarCab();" --> 
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

