<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : vercarga_objetivossmart_eva_00.asp
'Objetivo : ABM de objetivos smart de evaluacion
'Fecha	  : 16-05-2004
'Autor	  : CCRossi
'Modificacion: 28-07-2004 CCRossi cambiar el tema deel puntaje. ahora unico de evacab.
'...
'Modificacion: 22-10-2004 CCRossi "modificar" por "Grabar"
'Modificacion: 22-10-2004 CCRossi poner "(mi borrador)" si cejemp=-1 ( o sea es ABN....)
'Modificacion: 22-10-2004 Agegar cantidad de obj de cada tipo
'Modificacion: 22-10-2004 Sacar alerrt de puntuacion que no coincide, ya que es READONLY esto
'Modificacion: 02-11-2004 CCRossi -  no grabar puntuacion, ya que es seccion readonly y no habra modificaciones.
'Modificacion: 04-11-2004 CCRossi -  si es Superevaluador ABN, entonces que vea los puntajes.
'			   26-01-2005 Leticia Amadio - 
'              13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			   24/05/2007 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
on error goto 0

 Dim l_rs, l_rs1 
 Dim l_sql 
 Dim l_filtro
 Dim l_orden

 Dim col
 if cformed = -1 then
	col=7
 else 
	col=6
 end if	

'locales
 dim l_super
 dim l_puntaje
 dim l_puntajemanual
 dim l_evacabnro 
 dim l_evatevnro 
 dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
 dim l_mostrar '1 o 0 si teine que mostrsr la observacion. 
 dim l_evatipobjnro ' tipo objetivo
 dim l_cantidad ' cantidad de objetivos de un tipo
 'dim l_evatevnro ' 
 dim l_verevaluacion ' 
 '  dim l_evacabnro Xxxxxxxxxxxx
 
'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 dim l_empleg
  
 l_evldrnro = request.querystring("evldrnro")
 l_evapernro = request.querystring("evapernro")


 l_empleg = Session("empleg")
 if trim(l_empleg)="" then
	l_empleg = Request.QueryString("empleg")
 end if	

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
 end if
'___________________________________________________________________________________
function PasarComaAPunto(valor)
	dim l_numero
	dim l_ubicacion
	dim l_entero
	dim l_decimal
	l_numero = trim(valor)
	l_ubicacion = InStr(l_numero, ",")
	if l_ubicacion > 1 then
		l_ubicacion = l_ubicacion  - 1
		l_entero = left(l_numero, l_ubicacion)
		l_ubicacion = l_ubicacion  + 1
		l_decimal = right(l_numero, (len(l_numero) - l_ubicacion))
    	l_numero = l_entero & "." & l_decimal
    	PasarComaAPunto = l_numero
    else
		PasarComaAPunto = valor
	end if
end function	


'buscar la evacab
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro, evatevnro, empleado.empleg  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evacabnro = l_rs("evacabnro")
	l_evatevnro = l_rs("evatevnro")
	l_evaluador = l_rs("empleg")
 end if
 l_rs.close
 set l_rs=nothing


'XXXXXXXXXXXXXXXXX
l_sql= " SELECT "  '  ya esta l_evacab...
' SELECT evacabnro FROM evacabnro INNER JOIN evadetevldor
' WHERE evadetevldor.evldrnro = l_evldrnro
' l_evacabnro= l_rs("evacabnro")

' SELECT evldorcargada, evaseccnro FROM evadetevldrnro
' INNER JOIN vistos ON vistos.evldrnro= evadetevldor.evldrnro
' WHERE evacabnro= l_evacabnro AND evatevnro <> cautoevaluador AND evatevnro <> cevaluador 
'  aca obtengos si esta cargada o no la seccion (e realida no necesitaria la evaseccnro?)
 
 '  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXx evatevnro es el del evldrnro y no de quien se loguea!!!
'  ver si el evaluador termino la sección  ________________________________
'response.write l_evldrnro & " --"
'response.write l_evatevnro
'response.write " -  " & cautoevaluador
			' -----------------------------
'if l_evatevnro = cautoevaluador then 
	 
	 l_evacabnro = 0
	 		' ya está...........
	 'Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	 'l_sql = " SELECT evacabnro "
	 'l_sql = l_sql & " FROM evacabnro "
	 'l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
	 'l_sql = l_sql & " WHERE evadetevldor.evldrnro = l_evldrnro "
	  'if not l_rs1.eof then 
	 	'l_evacabnro = l_rs1("evacabnro")
	 'end if
	 
	 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	 
	 l_sql = " SELECT evldorcargada, evaseccnro "
	 l_sql = l_sql & " FROM evadetevldor "
	 l_sql = l_sql & " INNER JOIN evavistos ON evavistos.evldrnro= evadetevldor.evldrnro "
	 l_sql = l_sql & " WHERE evacabnro= " &  l_evacabnro 
	 l_sql = l_sql & " AND evatevnro <> " & cautoevaluador & " AND evatevnro <> " & cevaluador 
	 rsOpen l_rs1, cn, l_sql, 0 
  	 
	 l_verevaluacion = 0
 	 
	 if not l_rs1.EOF then 
 		if l_rs1("evldorcargada")= -1 then
 			l_verevaluacion = -1
		end if
	 end if
 	 
	 l_rs1.close
	 set l_rs1=nothing
'else 
	'l_verevaluacion = -1
'end if
 
 
'buscar puntaje cargado
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT puntaje, puntajemanual  "
 l_sql = l_sql & " FROM  evacab "
 l_sql = l_sql & " WHERE evacabnro   = " & l_evacabnro
  rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_puntajemanual= l_rs("puntajemanual")
	l_puntaje	   = l_rs("puntaje")
 end if
 l_rs.close
 set l_rs=nothing

 if trim(l_puntajemanual)<>"" then
	l_puntajemanual = PasarComaAPunto(l_puntajemanual)
 else	
	l_puntajemanual = ""
 end if	
 
 l_mostrar = "0"
' Response.Write l_empleg & "<br>" & l_evaluador
 if trim(l_empleg)<>"" and not isNull(l_empleg) then
	if trim(l_empleg) = trim(l_evaluador) then
		l_mostrar = "1"
	else	
		l_mostrar = "0"
	end if
 else
 	l_mostrar = "1"
 end if

' si es SUPEREVLAUDOR lo dejo ver
 l_super="0"
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evatevnro  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
 l_sql = l_sql & "	      AND evacab.evacabnro =  " & l_evacabnro
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
 l_sql = l_sql & "		  AND empleado.empleg   = " & l_empleg 
 l_sql = l_sql & " WHERE evatevnro<> " & cautoevaluador
 l_sql = l_sql & "   AND evatevnro<> " & cevaluador
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_super="1"
 end if
 l_rs.close
 set l_rs=nothing

 
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>

<script>

var entrada = new Array;
var salida = new Array;

entrada[0]=50;
entrada[1]=60;
entrada[2]=70;
entrada[3]=80;
entrada[4]=90;
entrada[5]=99;
entrada[6]=110;
entrada[7]=120;
entrada[8]=130;
entrada[9]=140;

salida[0]=0;
salida[1]=0.5;
salida[2]=1;
salida[3]=1.5;
salida[4]=2;
salida[5]=2.5;
salida[6]=3;
salida[7]=3.5;
salida[8]=4;
salida[9]=4.5;
salida[10]=5;

function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function ValidarDatos(ponde,alcanza)
{
	if (ponde.value=="") 
	{
		alert('Ingrese una Ponderación.');
		ponde.focus();
		return false;
	}	
	else
	if (isNaN(ponde.value) )
	{
		alert('Ingrese una Ponderación válida.');
		ponde.focus();
		return false;
	}	
	else
	if (alcanza.value=="") 
	{
		alert('Ingrese un Porcentaje Alcanzado.');
		alcanza.focus();
		return false;
	}	
	else
	if (isNaN(alcanza.value))
	{
		alert('Ingrese un Porcentaje Alcanzado válido.');
		alcanza.focus();
		return false;
	}	
	else
		return true;
		
}

function Controlar(ponde,alcanza){
	if (alcanza=="")
		alcanza.value=0;
		
	if (ponde.value>100)
	{
		alert('La Ponderación debe ser menor o igual a 100.');
		ponde.focus();
		return false;
	}	
	else
	{	
		var formElements = document.datos.elements;
		var i=1;
		var total=0;
		i=2;
		while (i<formElements.length-4)
		{
			total = Number(total) + Number(formElements[i].value);
			i = i + 7;
		}
		document.datos.totalponderacion.value=total;
		
		if ((Number(total)!==100))
		{
			alert('El total de Ponderación debe ser 100.');
			if ((Number(total) < 100) && (total!==""))
				return true;
			else
			{
				ponde.focus();
				return false;
			}	
		}
		else
			return true;	
	}
}	

function Sumar(){
	var formElements = document.datos.elements;
	var i=1;
	var total=0;
	i=2;
	while (i<formElements.length-3)
	{
		total = Number(total) + Number(formElements[i].value);
		i = i +7;
	}
	document.datos.totalponderacion.value=total;
		
	if ((total > 100) || (total < 100))
		alert('El total de Ponderación debe ser 100.');
}	

function Puntuacion(){
	var formElements = document.datos.elements;
	var i=0;
	var total=0;
	
	i=1;
	while (i<formElements.length-3)
	{
		
		formElements[i+3].value= formElements[i+1].value * formElements[i+2].value / 100;
		total = Number(total) + Number(formElements[i+3].value); 
		i = i + 7;
	}
	
	document.datos.totalpuntuacion.value=total;
	i=0;
	document.datos.puntuacion.value="";
	while (i<10)
	{
		if (Number(entrada[i])>Number(total))
		{
			document.datos.puntuacion.value=salida[i];
			i = 9;
		}	
		i=i + 1;
	}	
	
	if ((i>9) && (document.datos.puntuacion.value==""))
		document.datos.puntuacion.value=salida[10];

	// grabar valor en avadetevldor
	//var r = showModalDialog('grabar_puntuacion_eva_00.asp?evldrnro=<%=l_evldrnro%>&puntaje='+document.datos.puntuacion.value, '','dialogWidth:20;dialogHeight:20'); 
	<%if l_puntajemanual<>"" then%>	
	if (document.datos.puntajemanual.value=="")
		document.datos.puntajemanual.value=document.datos.puntuacion.value;
	if ((document.datos.puntuacion.value > (Number(document.datos.puntajemanual.value)+0.5))
		|| (document.datos.puntuacion.value < (Number(document.datos.puntajemanual.value)-0.5)))
		{
		<%if l_mostrar="1" then%>
		//alert('Deberá cambiar la Puntuación ingresada en Evaluación General.');
		<%end if%>
		divmanual.innerHTML='<font size=2 color=red><b>(<%=l_puntajemanual%>)</b></font>';
		}
	else	
		divmanual.innerHTML='<font size=2 color=black><b>(<%=l_puntajemanual%>)</b></font>';
	<%end if%>	
}	
		

function Validar(fecha)
{
	if (fecha == "") {
		alert("Debe ingresar la fecha .");
		return false;
		}
	else
		{
		return true;
		}
}
 	   

var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
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

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="Sumar();Puntuacion();">
<table>
    <tr>
        <th align=center colspan=2 class="th2">Objetivos</th>
        <%if cformed=-1 then%>
        <th align=center class="th2">Forma de Medici&oacute;n</th>
        <%end if%>
        <th align=center class="th2">Ponderaci&oacute;n</th>
        <th align=center class="th2">% alcanzado al final del per&iacute;odo</th>
        <th align=center class="th2">Puntuaci&oacute;n ponderada</th>
        <th align=center class="th2">Observaciones <%if cejemplo=-1 then%>(mi borrador)<%end if%></th>
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro,evaperfijo, evapernroeva, evaobjdext,evaobjformed, evaluaobj.evldrnro, evaobjpond, evaobjalcanza ,"
l_sql = l_sql & " evaobjsgto.evasgtotexto, "
l_sql = l_sql & " evaobjetivo.evatipobjnro, evatipobjdabr ,evatipobjorden "
l_sql = l_sql & " FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj  ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjetivo.evatipobjnro"
l_sql = l_sql & " INNER JOIN evaobjsgto ON evaobjsgto.evaobjnro = evaobjetivo.evaobjnro "
l_sql = l_sql & "        AND evaobjsgto.evldrnro = evaluaobj.evldrnro "
l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro
l_sql = l_sql & " ORDER BY evatipoobj.evatipobjorden "
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
l_evatipobjnro=""
l_cantidad = 0
do until l_rs.eof

	if l_evatipobjnro <> l_rs("evatipobjnro") then
		l_evatipobjnro= l_rs("evatipobjnro") 
		l_cantidad = 1 %>
		<tr>
        <td colspan="<%=col + 1%>"><b><%=l_rs("evatipobjdabr")%></b>
		</td>
		</tr>
	<%else
		l_cantidad = l_cantidad + 1
	end if%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("evaobjnro")%>)">
        <td align=center valign=top style="width:10">
			<b><%=l_cantidad%></b>
		</td>
		<td align=center valign=top>
        	<textarea readonly disabled name="evaobjdext<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjdext"))%></textarea>
		</td>
		<%if cformed=-1 then%>
        <td align=center>
			<textarea readonly disabled  name="evaobjformed<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=15 rows=5><%=trim(l_rs("evaobjformed"))%></textarea>
		</td>
		<%else%>
		<input type="hidden" name="evaobjformed<%=l_rs("evaobjnro")%>">
		<%end if%>

        <td align=center>
			<input readonly disabled  pond="<%=l_rs("evaobjnro")%>" type="text" name="evaobjpond<%=l_rs("evaobjnro")%>" size=5 value="<%=l_rs("evaobjpond")%>" >
		</td>
        <td align=center>
        <%if  l_mostrar="1" or l_super="1" or l_verevaluacion = -1 then   ' %>
			<input readonly disabled  type="text" name="evaobjalcanza<%=l_rs("evaobjnro")%>" size=5 value="<%=l_rs("evaobjalcanza")%>">
		<%else%>
			<input type="hidden" name="evaobjalcanza<%=l_rs("evaobjnro")%>" size=5 value="<%=l_rs("evaobjalcanza")%>">
			No habilitado para ver Porcentaje.
		<%end if%>
		</td>
        <td align=center>
        <%if  l_mostrar="1" or l_super="1"  or l_verevaluacion = -1 then   '%>
			<input readonly disabled type="text" class="blanc" name="puntacion<%=l_rs("evaobjnro")%>" size=4 >
		<%else%>
			<input readonly type="hidden" class="blanc" name="puntacion<%=l_rs("evaobjnro")%>" size=4>
			No habilitado para ver Puntuación.
		<%end if%>
		</td>
        <td align=center>
			<%if l_mostrar="1" or l_verevaluacion = -1 then   '%>
			<textarea readonly disabled name="evasgtotexto<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evasgtotexto"))%></textarea>
			<%else%>
			<input type=hidden name="evasgtotexto<%=l_rs("evaobjnro")%>">
			No habilitado para ver Observación.
			<%end if%>
			<input type="hidden" readonly disabled name="grabado<%=l_rs("evaobjnro")%>" size="1">
		</td>
    </tr>
<%
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
	<!-- t o t a  l e s ----------------------------------->
    <tr>
        <%if cformed=-1 then%>
			<td align=center>&nbsp;</td>
		<%end if%>
		<td align=center>&nbsp;</td>
        <td align=right><b>Totales</b></td>
		<td align=center><input style="background : #e0e0de;" readonly class="blanc" type="text" name="totalponderacion" size=5></td>
		<td align=center></td>
        <td align=center>
        <%if l_mostrar="1" or l_super="1" or l_verevaluacion = -1 then  '  %>
			<input style="background : #e0e0de;" readonly class="blanc" type="text" name="totalpuntuacion" size=5>
		<%else%>
			<input style="background : #e0e0de;" readonly class="blanc" type="hidden" name="totalpuntuacion" size=5>
			No habilitado
        <%end if%>
        </td>
		<td colspan=<%=col-4%>></td>
    </tr>
    <tr>
		<td align=center></td>
		<td align=center></td>
		<%if cformed=-1 then%>
        <td align=center></td>
        <%end if%>
        <td align=right colspan="<%=col-4%>"><b>Puntuaci&oacute;n de Objetivos (1 a 5)</b>
               	<input type="hidden" name="puntajemanual" value="<%=l_puntajemanual%>">
        </td>
		<td align=center>
	        <%if  l_mostrar="1" or l_super="1" or l_verevaluacion = -1 then  '   %>
			<input style="background : #e0e0de;" readonly class="blanc" type="text" name="puntuacion" size=5>
			<%else%>
			<input style="background : #e0e0de;" readonly class="blanc" type="hidden" name="puntuacion" size=5>
			No Habilitado
			<%end if%>
		</td>
		<td valign=top colspan="<%=col-4%>">&nbsp;</td>
		
    </tr>
</table>
<input type="Hidden" name="cabnro" value="0">
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</form>
</body>
</html>
