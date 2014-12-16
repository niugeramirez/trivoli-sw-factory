<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
on error goto 0

'Modificacion: el manejo de logeadoempleg...
'Modificado	: 03-02-2005 CCRossi. Cambiar Rotulo de columna para Codelco
' 			: 12-08-2005 - L.A. - tener en cuenta en la ult seccion cdo Auto=Gerente , Rev=Spcio y Rev=Gerente
'             -10-2005 - Leticia Amadio -  Adecuacion a Autogestion	

dim l_letra
dim l_pantalla
l_pantalla = request("pantalla")
if trim(l_pantalla) = "1024" then
	l_letra="style=font-size:8pt font-type:tahoma"
else	
	l_letra="style=font-size:7pt font-type:arial"
end if


dim l_logeadoempleg
l_logeadoempleg = Request.QueryString("logeadoempleg")

'response.write "<script>alert('form 02 : "&l_logeadoempleg&"');</script>"	

Dim l_rs
dim l_rs1
dim l_rs2
dim l_rs3
Dim l_rs_oblig
Dim l_cm

dim l_evldrnro ' para guardar el evldrnro de la otra seccion de objetivos
dim l_otrocab
Dim l_evapernro
Dim l_evaseccmail ' indica si al terminar la seccion/habilitarla avisa por mail 

Dim l_evatevnro
Dim l_evaluador
Dim l_sql

Dim l_empleado
Dim l_evaseccnro
Dim l_evacabnro
Dim l_revisor
dim l_revisorternro
Dim l_primera
Dim l_objetivo
Dim l_evarolaspdet
Dim l_habilitado
dim l_tipsecobj
dim l_etaprogcarga 
dim l_etaprogread 
	
  dim l_hora 
  dim l_arrhr
  dim l_fechahab 
  dim l_horahab   

dim l_logeado  
dim l_empleg

l_empleado	= request.querystring("ternro")
l_evaseccnro= request.querystring("evaseccnro")
l_evacabnro	= request.querystring("evacabnro")
l_revisor	= request.querystring("revisor")

function strto2(cad)
	if trim(cad) <>"" then
		if len(cad)<2 then
			strto2= "0" & cad
		else
			strto2= cad
		end if 
	else
		strto2= "00"
	end if	
end function

if trim(l_revisor) <>"" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT empleado.ternro FROM empleado WHERE empleado.empleg = " & l_revisor
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_revisorternro= l_rs("ternro")
	end if 
	l_rs.Close
	set l_rs=nothing
end if	


' chequear tipo de seccion
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "select tipsecobj from evasecc INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro WHERE evasecc.evaseccnro = " & l_evaseccnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_tipsecobj = l_rs("tipsecobj")
end if
l_rs.close
set l_rs=nothing

'cargar evaluadorse si no los tiene
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evadetevldor.evldrnro, evacabnro, evaluador, evarolaspdet,evatevdesabr , evadetevldor.evatevnro FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro= evadetevldor.evatevnro LEFT  JOIN evarolasp ON evarolasp.evarolnro = evatipevalua.evarolnro "
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro & " AND   evaseccnro = " & l_evaseccnro
rsOpen l_rs1, cn, l_sql, 0
l_evarolaspdet=""

do while not l_rs1.eof

	if isnull(l_rs1("evaluador")) or trim(l_rs1("evaluador"))="" then

		Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evaluador FROM evadetevldor WHERE evacabnro = " & l_evacabnro
		l_sql = l_sql & " AND   evatevnro = " & l_rs1("evatevnro") & " AND   evldrnro <> " & l_rs1("evldrnro")
		rsOpen l_rs2, cn, l_sql, 0
		if not l_rs2.eof then
				if trim(l_rs2("evaluador"))<>"" and not isnull(l_rs2("evaluador")) then
				set l_cm = Server.CreateObject("ADODB.Command")
				l_sql = "UPDATE evadetevldor SET evaluador = " & l_rs2("evaluador") & " WHERE evldrnro = " & l_rs1("evldrnro")
				l_cm.activeconnection = Cn
				l_cm.CommandText = l_sql
				cmExecute l_cm, l_sql, 0
				else
				Response.Write("<script>alert('No se asignó el "&l_rs1("evatevdesabr")&".\nEsto puede deberse a que no exista el programa de asignación, o bien el programa de asignación no encontró un empleado.\n\n Ingrese en Relacionar Empleados, Evaluadores y asígnelo manualmente.');</script>")
				Response.Write("<script>parent.close();</script>")
				Response.End
				end if
		end if
		l_rs2.close
		set l_rs2=nothing	
	end if
	l_rs1.MoveNext
loop
l_rs1.close
set l_rs1=nothing

if trim(l_evarolaspdet)<>"" then%>
	<script>
	window.location.reload();
	</script>
<%end if%>
<%
' ______________________________________________________
' me fijo si Autoev = Gerente , REv=Socio o Rev=Socio   
if cint(cdeloitte)=-1 then 
' dim primero
dim l_evaluador_auto, l_evaluador_rev ', l_evaluador1, l_evaluador2
dim l_evatevnro1, l_evatevnro2

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")

l_evatevnro1 = 0
l_evatevnro2 = 0
l_evaluador_rev=0
l_evaluador_auto=0

l_sql = "SELECT  ultimasecc FROM evasecc WHERE evasecc.evaseccnro =" & l_evaseccnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then 
	if l_rs("ultimasecc") = -1 then 
		l_sql = "SELECT evaluador, evatevnro  FROM evadetevldor  WHERE evacabnro="& l_evacabnro & " AND evaseccnro="& l_evaseccnro
		rsOpen l_rs1, cn, l_sql, 0
		do while not l_rs1.EOF 
			if l_rs1("evatevnro")= cautoevaluador then
				l_evaluador_auto = l_rs1("evaluador") 
			end if 
			if l_rs1("evatevnro")= cevaluador then 
				l_evaluador_rev = l_rs1("evaluador") 
			end if 
		l_rs1.MoveNext 
		loop 
		l_rs1.close 
		
		' --------------------------------------------------------
		'if l_evaluador_auto <> 0 then
		l_sql = " SELECT evaluador, evatevnro  FROM evadetevldor "
		l_sql = l_sql & " WHERE  evacabnro="& l_evacabnro & " AND evaseccnro="& l_evaseccnro
		l_sql = l_sql & "		AND evatevnro <> " & cautoevaluador & " AND evatevnro <> "& cevaluador
		l_sql = l_sql & "		AND evaluador="& l_evaluador_auto
		rsOpen l_rs1, cn, l_sql, 0
		if not l_rs1.eof then
			l_evatevnro1 = l_rs1("evatevnro")
		end if
		
		l_rs1.close 
		
		l_sql = " SELECT evaluador, evatevnro  FROM evadetevldor "
		l_sql = l_sql & " WHERE  evacabnro="& l_evacabnro & " AND evaseccnro="& l_evaseccnro
		l_sql = l_sql & "		AND evatevnro <> " & cautoevaluador & " AND evatevnro <> "& cevaluador
		l_sql = l_sql & "		AND evaluador="& l_evaluador_rev
		l_sql = l_sql & " ORDER BY evatevnro DESC"
		rsOpen l_rs1, cn, l_sql, 0
		if not l_rs1.eof then
			l_evatevnro2 = l_rs1("evatevnro")
		end if
		
		
		l_rs1.close 
		set l_rs1 = nothing 
		
		
	end if 
end if  

l_rs.close 
set l_rs=nothing 
	
end if
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Secciones del Formulario - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>
var jsSelRow = null;
function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function cambiarevaluador(evldrnro){
	var nuevo="";
	abrirVentana('form_carga_eva_04.asp?evldrnro='+evldrnro,'',450,100);	
}

function actualizar(){
parent.recargar();
}

function Deseleccionar(fila)
{
 document.all.primera.className = "MouseOutRow";
 if (jsSelRow != null)
	 fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,evldrnro,habilitado,etaprogcarga,etaprogread,aprobada,logeado)
{
 Deseleccionar(jsSelRow);
 
 document.datos.cabnro.value = cabnro;

 parent.actualizarcarga(evldrnro,<%= l_evaseccnro %>,habilitado,etaprogcarga,etaprogread,aprobada,logeado);
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th><font <%=l_letra%>><%if ccodelco=-1 then%>Num.<%else%>Emp.<%end if%></th>
        <th><font <%=l_letra%>>Apellido y Nombre</th>
		<th><font <%=l_letra%>><%if ccodelco=-1 then%>Actor<%else%>Tipo<%end if%></th>
    </tr>
<%
'============================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")	
l_sql = "SELECT DISTINCT empleado.ternro, empleado.empleg, empleado.terape, empleado.ternom, habilitado, evatipevalua.evatevdesabr, evadetevldor.evldrnro, evadetevldor.evatevnro, evaseceta.etaprogcarga, evaseceta.etaprogread, evaoblieva.evaobliorden, evacab.cabaprobada FROM evacab "
l_sql = l_sql & " inner join evadetevldor on evacab.evacabnro= evadetevldor.evacabnro left join empleado on evadetevldor.evaluador= empleado.ternro inner join evatipevalua on evadetevldor.evatevnro= evatipevalua.evatevnro "
l_sql = l_sql & " inner join evasecc on evadetevldor.evaseccnro= evasecc.evaseccnro inner join evaoblieva on evaoblieva.evatevnro= evadetevldor.evatevnro and evaoblieva.evaseccnro=evasecc.evaseccnro left  join evaseceta on evaseceta.evaseccnro= evadetevldor.evaseccnro "
l_sql = l_sql & " AND  evaseceta.evatipnro= evasecc.evatipnro AND  evaseceta.evaetanro= evacab.evaetanro "
l_sql = l_sql & " WHERE evacab.empleado = " & l_empleado & " AND evacab.evacabnro = " & l_evacabnro & " AND evadetevldor.evaseccnro=" & l_evaseccnro & " ORDER BY evaoblieva.evaobliorden " 
'response.write l_sql & "<br>"
rsOpen l_rs, cn, l_sql, 0 
l_primera= true
 
Dim objOpenFile, objFSO, strPath
	
do until l_rs.eof
	'response.write l_rs("etaprogcarga") & " - "& l_rs("etaprogread") & "<br>"
	
	if trim(l_rs("etaprogcarga"))="" or isnull(l_rs("etaprogcarga")) then
		l_etaprogcarga = ""
	else
		l_etaprogcarga = l_rs("etaprogcarga")
		strPath = Server.MapPath(l_etaprogcarga)
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If not objFSO.FileExists(strPath) Then
		   l_etaprogcarga="*"
		End If
	end if
	if trim(l_rs("etaprogread"))="" or isnull(l_rs("etaprogread")) then
		l_etaprogread = ""
	else	
		l_etaprogread = l_rs("etaprogread")
		strPath = Server.MapPath(l_etaprogread)
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If not objFSO.FileExists(strPath) Then
		   l_etaprogread="*"
		End If
	end if

	'response.write "<script>alert('form 02 : "&l_logeadoempleg&"');</script>"	
	'response.write "<script>alert('form 02-empleg actual : "&l_rs("empleg")&"');</script>"	

if trim(l_logeadoempleg)<>"" then
	if cdbl(l_logeadoempleg)<>cdbl(l_rs("empleg")) then ' el evluador no es el logeado!!!
		l_logeado=0
	else	
		l_logeado=-1
	end if
else
	l_logeado=-1
end if	
	'response.write "<script>alert('LOGEADO : "&l_logeado&"');</script>"	
if l_primera then
	response.write "<script>parent.actualizarcarga("&l_rs("evldrnro")&","&l_evaseccnro&","&l_rs("habilitado")&",'"&l_etaprogcarga&"','"&l_etaprogread &"',"&l_rs("cabaprobada")&","&l_logeado&");</script>"
end if	

if cint(cdeloitte)= -1 then 
	if l_rs("evatevnro") <> l_evatevnro1 and l_rs("evatevnro") <> l_evatevnro2  then 
%>
    <tr <%if l_primera then%> id="primera" class="SelectedRow"<%end if%>  onclick="Javascript:Seleccionar(this,'<%= l_rs("evldrnro")%>',<%= l_rs("evldrnro")%>,<%= l_rs("habilitado")%>,'<%= l_etaprogcarga%>','<%= l_etaprogread%>',<%= l_rs("cabaprobada")%>,<%=l_logeado%>)">
        <td><font <%=l_letra%>><%= l_rs("empleg")%></td>
        <td nowrap><font <%=l_letra%>><%= l_rs("terape") & ", " & l_rs("ternom")%></td>
        <td nowrap><font <%=l_letra%>><%= l_rs("evatevdesabr")%></td>
    </tr>
<%  	l_primera = false 
	end if 
else 
%>
    <tr <%if l_primera then%> id="primera" class="SelectedRow"<%end if%>  onclick="Javascript:Seleccionar(this,'<%= l_rs("evldrnro")%>',<%= l_rs("evldrnro")%>,<%= l_rs("habilitado")%>,'<%=l_etaprogcarga%>','<%=l_etaprogread%>',<%= l_rs("cabaprobada")%>,<%=l_logeado%>)">
        <td><font <%=l_letra%>><%= l_rs("empleg")%></td>
        <td nowrap><font <%=l_letra%>><%= l_rs("terape") & ", " & l_rs("ternom")%></td>
        <td nowrap><font <%=l_letra%>><%= l_rs("evatevdesabr")%></td>
    </tr>
<%	
	l_primera = false
end if


l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="pgrrest" value="">
</form>
</body>
</html>
