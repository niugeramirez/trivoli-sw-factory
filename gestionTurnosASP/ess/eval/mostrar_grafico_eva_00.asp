<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'____________________________________________________________________________________
'Archivo  : mostrar_grafico_eva_00.asp
'Objetivo : 
'Fecha	  : 
'Autor	  :
'Modificado: 21-09-2004: CCRossi. Cambiar forma de buscar evaluadores, para que no quede hardcode
' 			 sacando constantes de const_eva.
' 			 30-11-2004-CCRossi. Dar vuelta los ejes.(Pedido Accor)
'			 __- 4- 2006- LPA. - cambiar la funcio cdbl por round, cuando se arma el grafico (con cdbl daba mal)
'				 agregar titulo al grafico 
'			 Contar Cantidad de Competeencias Evaluadas por el Potencial y no usar la cant del evaluador - varios cambio para los calculos
' 			 11-07-2006 - lA - adecuación Autogestion	
'			   -09-2006 -  Aplicar funciones clng() y cdbl() en algunas variables
'--------------------------------------------------------------------------------------
on error goto 0
' Variables
' de parametros entrada
  
' de uso local  

  Dim l_suma
  
  Dim l_sumad
  Dim l_maximo
  Dim l_producto
  Dim l_productop
   
  Dim l_cantidad
  Dim l_cantidadp
  Dim l_sumap
  
  Dim l_porcentaje
  Dim l_porcentajep
  Dim l_porcentajed
  Dim l_puntajefinal
  Dim l_evaluador
  Dim l_potencial
  Dim l_auto
  Dim l_evaseccnro
  Dim l_evatevnro
  Dim l_usarporcen

  Dim i
  Dim j
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  Dim l_empleado
  Dim l_esxestr
  l_esxestr="NO"
' parametros de entrada---------------------------------------  
  l_evldrnro = Request.QueryString("evldrnro")
  l_empleado = Request.QueryString("empleado")
%>


<!--#include virtual="/serviciolocal/shared/inc/mostrar_totales_00.inc"--> 
<!-- En el include se RENOMBRA el EVLDRNRO con el evldrnro correspondiente a la sección de Competencias -->
  
<%
if l_error = 0 then
	l_producto=0
	l_productop=0
	
	'Competencias evaluadas por evaluador
	if Clng(l_cantidad) <> 0 then 
	  l_producto     = cdbl(l_maximo) * cdbl(l_cantidad)
	end if 
	
	if l_producto <> 0 then
		if (l_esxestr="SI" and cint(l_usarporcen)=-1) then
			l_porcentajed   = (cdbl(l_suma)  / cdbl(l_maximo) )  * 100
		else
		  	l_porcentajed   = cdbl(l_suma)  / cdbl(l_producto)  * 100
	 	end if
	else 
		  l_porcentajed   = 0
	end if
	
	
	'Competencias evaluadas por Potencial
	if cdbl(l_cantidadp) <> 0 then
	  l_productop     = cdbl(l_maximo) * cdbl(l_cantidadp)
	end if 
	
	'if l_producto <> 0 then
	 ' l_porcentajep   = cdbl(l_suma)  / cdbl(l_producto)  * 100
	'else 
	 ' l_porcentajep   = 0
	'end ifº
	
'response.write "l_maximo= "  & l_maximo & "<br>"
'response.write "l_producto= "  & l_producto & "<br>"
'response.write "l_suma= "  & l_suma & "<br>"
'response.write "l_porcentajed= "  & l_porcentajed & "<br>"
%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>An&aacute;lisis de Calificaciones</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table width="50" align="center">
<tr>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td colspan="21" align="center"> &nbsp; <br>
	<b>GR&Aacute;FICO DESEMPEÑO/POTENCIAL DE COMPETENCIAS EVALUADAS</b><br>	&nbsp;
	</td>
</tr>
<tr>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td colspan="21" class="grafico" align="center">&nbsp;</td>
</tr>
<%
' ======================================================================================
 ' SUMA TOTAL DE EV. POTENCIAL  O COMP. POTENC. 										
  if l_esxestr="SI" then
    	if cint(l_usarporcen)=-1 then    ' Suma de resultados ponderados
			 l_sql = " SELECT SUM(evarestot) AS suma "
			 l_sql = l_sql & " FROM  evaresultado "
			 l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
			 l_sql = l_sql & " INNER JOIN evafactor		ON evafactor.evafacnro = evaresultado.evafacnro "
			 l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
			 l_sql = l_sql & " WHERE evaresultado.evldrnro = " & l_potencial
			 rsOpen l_rs, cn, l_sql, 0
			 'Response.Write l_sql
			 if not l_rs.eof then 
			   	l_sumap = l_rs("suma")
			 else
			   Response.Write("<script>alert('Se necesitan valores Potenciales para realizar el gráfico.')</script>")
			   Response.End
			 end if
			 l_rs.Close
		else
	  		 l_sql = " SELECT SUM(evatrvalor) AS suma "
			 l_sql = l_sql & " FROM  evaresultado "
			 l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
			 l_sql = l_sql & " INNER JOIN evafactor	    ON evaresultado.evafacnro = evafactor.evafacnro "
			 l_sql = l_sql & " INNER JOIN evadescomp    ON evadescomp.evafacnro=evafactor.evafacnro  "
			 l_sql = l_sql & " AND evadescomp.estrnro IN ("& l_estrnros & ")"
			 l_sql = l_sql & " INNER JOIN evadetevldor  ON evadetevldor.evldrnro = evaresultado.evldrnro "
			 l_sql = l_sql & " WHERE evaresultado.evldrnro = " & l_potencial
			 rsOpen l_rs, cn, l_sql, 0
			' Response.Write l_sql
			 if not l_rs.eof then 
			   l_sumap = l_rs("suma")
			 else
			   Response.Write("<script>alert('Se necesitan valores Potenciales para realizar el gráfico.')</script>")
			   Response.End
			 end if
			 l_rs.Close
		end if  ' de si usa procentaje de ponderacion o no.
  else
  	 ' Primero busco busco evaluador potencial
	  l_sql = " SELECT SUM(evatrvalor) AS suma "
	  l_sql = l_sql & " FROM  evaresultado "
	  l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
	  l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
	  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =  " & l_evaseccnro 
	  l_sql = l_sql & " AND   evaresultado.evldrnro = " & l_potencial
	  rsOpen l_rs, cn, l_sql, 0
	  'response.write l_Sql
	  if not l_rs.eof and l_potencial <> 0 then 
			l_sumap = l_rs("suma")
	  else	
			  l_rs.close  
	  		' si no hay evaluador potencial busco si se definio una competencia potencial
			  l_sql = " SELECT evatrvalor "
			  l_sql = l_sql & " FROM  evaresultado "
			  l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
			  l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
			  l_sql = l_sql & " INNER JOIN evafactor	 ON evaseccfactor.evafacnro = evafactor.evafacnro "
			  l_sql = l_sql & "		 AND   evafactor.evafacpot = -1 " 
			  l_sql = l_sql & " INNER JOIN evadetevldor	 ON evadetevldor.evldrnro = evaresultado.evldrnro "
			  l_sql = l_sql & "		 AND   evadetevldor.evacabnro =  " & l_evacabnro
			  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =  " & l_evaseccnro  & " AND evaresultado.evldrnro=" & l_evldrnro ' el evldrnro se renombro en el inc
			  	 ' response.write l_evldrnro
			  'Response.Write l_sql
			  rsOpen l_rs, cn, l_sql, 0
			  if not l_rs.eof then 
					l_sumap = cdbl(l_rs("evatrvalor")) '* l_cantidad 
			  else	
					Response.Write("<script>alert('Se necesitan valores Potenciales para realizar el gráfico.')</script>")
					Response.End
			  end if
			  l_rs.Close
	  end if
end if ' de chequeo de si es por Estr o NO

if trim(l_sumap)="" or isnull(l_sumap) then
  Response.Write("<script>alert('Se necesitan valores Potenciales para realizar el gráfico.')</script>")
  Response.End
end if

'response.write "sumap= " & l_sumap & "<br>"
l_porcentajep   = 0
if l_potencial <> 0 then ' el que evaluo es el rol Potencial
	if (l_esxestr = "SI") and cint(l_usarporcen)=-1 then
	  	if l_maximo <> "" then
		  	l_porcentajep   = ( cdbl(l_sumap)  /  cdbl(l_maximo) )  * 100
		end if
	else 
		if l_productop <> 0 then
			l_porcentajep   = cdbl(l_sumap)  / cdbl(l_productop)  * 100
		end if
	end if
else ' Si Rol potencial es 0, entonces es Competencia Potencial.
	  if l_maximo <> "" then
	  	l_porcentajep  = cdbl(l_sumap) / cdbl(l_maximo) * 100
	  end if
end if

'response.write "l_cantid= "  & l_cantidad & "<br>"
'response.write "l_maximo= "  & l_maximo & "<br>"
'response.write "l_sumap= "  & l_sumap & "<br>"
'response.write "l_porcentajed= "  & l_porcentajed & "<br>"
'response.write "porcentajep= " & l_porcentajep & "<br>"

' comienza el grafico ==================================================================
i = 0 
do while i <= 100%>
	<tr height="2">
		<%if i = 0 then%>
			<td rowspan="21" width="10" class="grafico"><B>Desempeño</B></td>
		<%end if %>
		<%if i = 0 then%>
			<td rowspan="7" width="10" class="grafico"><B>Alto</B></td>
		<%end if %>
		<%if i = 35 then%>
			<td rowspan="7" width="10" class="grafico"><B>Medio</B></td>
		<%end if %>
		<%if i = 70 then%>
			<td rowspan="7" width="10" class="grafico"><B>Bajo</B></td>
		<%end if %>
		
		<td width=2 align=center class="grafico"><%=(100 - i)%></td>
		<%
		  j = 0 
		  do while j <= 100
			if ( (round(l_porcentajed,2) >= (100-i)) and (round(l_porcentajed,2) > (100-i-5)) and ((round(l_porcentajed,2)-100 + i ) < 5) ) then
				if  round(l_porcentajep,2) >= j and (round(l_porcentajep,2) < (j+5)) then%>
					<td width="2" class="pgrilla">&nbsp;</td>
				<%else%>
					<td width="2" class="grilla">&nbsp;</td>
				<%end if
			else%>
				<td width="2" class="grilla">&nbsp;</td>
		   <%end if		
		j = j + 5
		loop%>
	</tr>
	<%i = i + 5
loop

j = 0
%>
<tr>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
    <%do while j <= 100%>
	   <td class="grafico"><%=j%></td>
	 <%j = j + 5
  loop%>
</tr>

<tr>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td colspan="7" class="grafico" align="center"><b>Bajo</b></td>
	<td colspan="7" class="grafico" align="center"><b>Medio</b></td>
	<td colspan="7" class="grafico" align="center"><b>Alto</b></td>
</tr>

<tr>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td colspan="21" class="grafico" align="center"><b>POTENCIAL</b></td>
</tr>
</table>
<%else%>
<table width="50" align="center">
<tr>
	<td colspan="21" class="grafico" align="center"><b>No hay un rol definido como Potencial.</b></td>
</tr>
</table>
<%end if%>

</body>
</html>
