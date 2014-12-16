<% 
function resumen_grafico
' de uso local  
  Dim l_suma
  Dim l_sumap
  Dim l_sumad
  Dim l_maximo
  Dim l_producto
  Dim l_cantidad
  Dim l_porcentaje
  Dim l_porcentajep
  Dim l_porcentajed
  Dim l_puntajefinal
  Dim l_evaluador
  Dim l_potencial
  Dim l_auto
  Dim l_evaseccnro
  Dim l_evatevnro

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
  
' selects a la base para obtener seccnro, los evldrnro segun  el tipo evaluador
' y calcular los valores



' Buscar el evaseccnro para tiposecnro= 2 (competencias)
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evadet.evaseccnro FROM evadetevldor "
l_sql = l_sql & "inner join evacab ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & "inner join evadet  ON evadet.evacabnro=evacab.evacabnro "
l_sql = l_sql & "inner join evasecc ON evasecc.evaseccnro=evadet.evaseccnro AND evasecc.tipsecnro = 2 "
l_sql = l_sql & "WHERE evadetevldor.evacabnro=" & l_evacabnro 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evaseccnro= l_rs("evaseccnro")
else	
	l_evaseccnro= "0"
end if	
l_rs.Close


'Buscar evaluador
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaluadores.evldrnro from evadetevldor "
l_sql = l_sql & "inner join evacab ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & "inner join evadetevldor evaluadores ON evacab.evacabnro = evaluadores.evacabnro "
l_sql = l_sql & "WHERE evadetevldor.evacabnro=" & l_evacabnro
l_sql = l_sql & " and evaluadores.evatevnro= " & cevaluador
l_sql = l_sql & " and evaluadores.evaseccnro=" & l_evaseccnro 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evaluador = l_rs("evldrnro")
else 
	l_evaluador = "0"	
end if	
l_rs.Close
' buscar potencial
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaluadores.evldrnro from evadetevldor inner join evacab ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & "inner join evadetevldor evaluadores ON evacab.evacabnro = evaluadores.evacabnro "
l_sql = l_sql & "WHERE evadetevldor.evacabnro=" & l_evacabnro 
l_sql = l_sql & " and evaluadores.evatevnro= " & cpotencial 
l_sql = l_sql & " and evaluadores.evaseccnro=" & l_evaseccnro 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_potencial = l_rs("evldrnro")
else 
	l_potencial = "0"	
end if	
l_rs.Close


	' buscar AUTO
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaluadores.evldrnro from evadetevldor inner join evacab ON evadetevldor.evacabnro=evacab.evacabnro "
	l_sql = l_sql & "inner join evadetevldor evaluadores ON evacab.evacabnro = evaluadores.evacabnro "
	l_sql = l_sql & "WHERE evadetevldor.evacabnro=" & l_evacabnro
	l_sql = l_sql & " and evaluadores.evatevnro= " & cautoevaluador
	l_sql = l_sql & " and evaluadores.evaseccnro=" & l_evaseccnro 
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_auto = l_rs("evldrnro")
	else 
		l_auto = "0"	
	end if	
	l_rs.Close

' Buscar el maximo valor    
  l_sql = "SELECT MAX(evatrvalor) AS maximo "
  l_sql = l_sql & " FROM  evatipresu "
  l_sql = l_sql & " INNER JOIN evaresu       ON evaresu.evatrnro = evatipresu.evatrnro "
  l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then 
	l_maximo = l_rs("maximo")
  end if
  l_rs.Close

' Buscar cantidad de competencias para la seccion  
  l_sql = "SELECT count(evafacnro) AS cantidad "
  l_sql = l_sql & " FROM  evaseccfactor "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then 
	l_cantidad = l_rs("cantidad")
  end if
  l_rs.Close


%>  
<table width="50" align="center" border="1" cellpadding="0" cellspacing="0">
<%
  
' para el grafico --------------------------------------------------------
' Calcular Suma Total POTENCIAL 
  l_sql = " SELECT SUM(evatrvalor) AS suma "
  l_sql = l_sql & " FROM  evaresultado "
  l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
  l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =  " & l_evaseccnro 
  l_sql = l_sql & " AND   evaresultado.evldrnro = " & l_potencial
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then 
	l_sumap = l_rs("suma")
  end if
  l_rs.Close

' Calcular Suma Total  DESEMPEÑO
  l_sql = " SELECT SUM(evatrvalor) AS suma "
  l_sql = l_sql & " FROM  evaresultado "
  l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
  l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =  " & l_evaseccnro 
  l_sql = l_sql & " AND   evaresultado.evldrnro = " & l_evaluador
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then 
	l_sumad = l_rs("suma")
  end if
  l_rs.Close

  l_porcentajep = cdbl(l_sumap)  / (cdbl(l_maximo) * cdbl(l_cantidad) ) * 100
  l_porcentajed = cdbl(l_sumad)  / (cdbl(l_maximo) * cdbl(l_cantidad) ) * 100

i = 0 
do while i <= 100%>
	<tr height="2">
		<%if i = 0 then%>
			<td rowspan="21" width="10" class="grafico"><B>Potencial</B></td>
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
if (cdbl(l_porcentajep) >= (100-i)) and (cdbl(l_porcentajep) > (100-i-5)) and ((cdbl(l_porcentajep)-100 + i ) < 5) then
if  cdbl(l_porcentajed) >= j and (cdbl(l_porcentajed) < (j+5)) then%>
								<td width="2">
									<img align="absmiddle" src="/turnos/shared/images/icon10.gif" border="0">
								</td>
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
</table>
<table width="50" align="center" border="0" cellpadding="0" cellspacing="0">
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
	<td colspan="21" class="grafico" align="center"><b>Desempeño</b></td>
</tr>
</table>
<%
end function
%>