<% 
'14-12-2004-CCRossi- cambiar l_rs por l_rsx ya que l_rs se usa en la asp que lo incluye.
'14-12-2004-CCRossi- Cambiar el mensaje de error cuando no se evaluo el potencial
'-----------------------------------------------------------------------------

function resumen_grafico
' de uso local  
  Dim l_suma
  Dim l_sumad
  Dim l_maximo
  Dim l_producto
  
  dim l_terape
   
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

  Dim i
  Dim j
  
' de base de datos  
  Dim l_sql
  Dim l_rsx
  Dim l_rs1
  Dim l_cm

  Dim l_evldrnro
  Dim l_empleado

  'l_evldrnro = Request.QueryString("evldrnro")
  'l_empleado = Request.QueryString("empleado")

dim l_Error
Dim l_tipseccalif
Dim l_rs_oblig

Set l_rsx = Server.CreateObject("ADODB.RecordSet")

l_Error = 0


'Busca el codigo del tipo de seccion competencias
l_sql = "SELECT evasecc.tipsecnro "
l_sql = l_sql & " FROM evasecc "
l_sql = l_sql & " INNER JOIN evaseccfactor ON evasecc.evaseccnro = evaseccfactor.evaseccnro "
l_sql = l_sql & " INNER JOIN evadetevldor  ON evasecc.evaseccnro  = evadetevldor.evaseccnro "
l_sql = l_sql & " WHERE evadetevldor.evacabnro=" & l_evacabnro
rsOpen l_rsx, cn, l_sql, 0 
if not l_rsx.eof then
	l_tipseccalif = l_rsx("tipsecnro")
else
	l_tipseccalif = "0"
end if
l_rsx.Close

' Buscar el evaseccnro para tiposecnro= l_tipseccalif (competencias)
l_sql = "SELECT evadet.evaseccnro "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab  ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evadet  ON evadet.evacabnro      =evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro    =evadet.evaseccnro "
l_sql = l_sql & "		AND evasecc.tipsecnro = "& l_tipseccalif
l_sql = l_sql & " WHERE evacab.evacabnro=" & l_evacabnro
rsOpen l_rsx, cn, l_sql, 0
if not l_rsx.eof then
	l_evaseccnro= l_rsx("evaseccnro")
end if	
l_rsx.Close

'Buscar tipo evaluador y el evldrnro de la seccion de competencias, para el eveluador actual
Set l_rsx = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaluador.evldrnro, evaluador.evatevnro "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab				    ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evadetevldor evaluador ON evacab.evacabnro = evaluador.evacabnro "
l_sql = l_sql & "	     AND evaluador.evatevnro = evadetevldor.evatevnro "
l_sql = l_sql & " WHERE evacab.evacabnro       = " & l_evacabnro
l_sql = l_sql & "   AND evaluador.evaseccnro   = " & l_evaseccnro 'seccion de comptencias
rsOpen l_rsx, cn, l_sql, 0
if not l_rsx.eof then
	l_evatevnro = l_rsx("evatevnro")
	l_evldrnro  = l_rsx("evldrnro")
end if	
l_rsx.Close

'Buscar tipo evaluador y el evldrnro de la seccion de competencias, para el eveluador actual
Set l_rsx = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT terape "
l_sql = l_sql & " FROM v_empleado "
l_sql = l_sql & " INNER JOIN evacab	ON evacab.empleado=v_empleado.ternro "
l_sql = l_sql & " WHERE evacab.evacabnro       = " & l_evacabnro
rsOpen l_rsx, cn, l_sql, 0
if not l_rsx.eof then
	l_terape = l_rsx("terape")
end if	
l_rsx.Close

'=========================================================================================
' buscar potencial
Set l_rsx = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaluadores.evldrnro "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab		  ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evadetevldor evaluadores ON evacab.evacabnro = evaluadores.evacabnro "
l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evaluadores.evatevnro "
l_sql = l_sql & "        AND evatipevalua.evatevpot = -1 " 
l_sql = l_sql & " WHERE evadetevldor.evldrnro=" & l_evldrnro 
l_sql = l_sql & "   and evaluadores.evaseccnro=" & l_evaseccnro 
rsOpen l_rsx, cn, l_sql, 0
if not l_rsx.eof then
	l_potencial = l_rsx("evldrnro")
else
	l_rsx.close
	set l_rsx=nothing
	Set l_rsx = Server.CreateObject("ADODB.RecordSet")
	l_sql = " SELECT evaresultado.evafacnro "
	l_sql = l_sql & " FROM  evaresultado "
	l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
	l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
	l_sql = l_sql & " INNER JOIN evafactor	 ON evaseccfactor.evafacnro = evafactor.evafacnro "
	l_sql = l_sql & "		 AND   evafactor.evafacpot = -1 " 
	l_sql = l_sql & " INNER JOIN evadetevldor	 ON evadetevldor.evldrnro = evaresultado.evldrnro "
	l_sql = l_sql & "		 AND   evadetevldor.evacabnro =  " &l_evacabnro
	l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =  " & l_evaseccnro 
	rsOpen l_rsx, cn, l_sql, 0
	if l_rsx.eof then 
		Response.Write("<script>alert('No se puede emitir el Gráfico de Desempeño para: "&l_terape&".\n\nVerifique que se haya evaluado el Portencial.\nTambién que se haya definido un Rol Potencial o Competencia Potencial.')</script>")
		l_Error = -1
	end if
	l_rsx.close
	set l_rsx=nothing
end if	


'=========================================================================================
' Calcular Suma Total
 Set l_rsx = Server.CreateObject("ADODB.RecordSet")
 l_sql = " SELECT SUM(evatrvalor) AS suma "
 l_sql = l_sql & " FROM  evaresultado "
 l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
 l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
 l_sql = l_sql & " INNER JOIN evafactor		ON evafactor.evafacnro = evaresultado.evafacnro "
 l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
 l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro 
 l_sql = l_sql & " AND   evaresultado.evldrnro    = " & l_evldrnro
 rsOpen l_rsx, cn, l_sql, 0
 if not l_rsx.eof then 
	l_suma = l_rsx("suma")
 end if
 l_rsx.Close
 'Buscar el maximo valor    
  l_sql = "SELECT MAX(evatrvalor) AS maximo "
  l_sql = l_sql & " FROM  evatipresu "
  l_sql = l_sql & " INNER JOIN evaresu       ON evaresu.evatrnro = evatipresu.evatrnro "
  l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
  Set l_rsx = Server.CreateObject("ADODB.RecordSet")
  rsOpen l_rsx, cn, l_sql, 0
  if not l_rsx.eof then 
	l_maximo = l_rsx("maximo")
  end if
  l_rsx.Close

'response.write l_maximo & "<br>"

  l_sql = "SELECT COUNT(evaseccfactor.evafacnro) AS cantidad "
  l_sql = l_sql & " FROM evaseccfactor "
  l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
  l_sql = l_sql & " AND EXISTS "
  l_sql = l_sql & " (SELECT * FROM evaresultado "
  l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresultado.evatrnro "
  l_sql = l_sql & " AND evatipresu.evatrnro IS NOT NULL "
  l_sql = l_sql & " WHERE evaresultado.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " AND   evaresultado.evldrnro = " & l_evldrnro
  l_sql = l_sql & ")"
  Set l_rsx = Server.CreateObject("ADODB.RecordSet")
  rsOpen l_rsx, cn, l_sql, 0
  if not l_rsx.eof then 
	l_cantidad = l_rsx("cantidad")
  end if
  l_rsx.Close
  
'response.write l_cantidad & "<br>"


if l_error = 0 then
	if l_cantidad <> 0 then
	  l_producto     = cdbl(l_maximo) * cdbl(l_cantidad)
	end if 

	if l_producto <> 0 then
	  l_porcentajed   = cdbl(l_suma)  / cdbl(l_producto)  * 100
	else 
	  l_porcentajed   = 0
	end if
'response.write "l_maximo= "  & l_maximo & "<br>"
'response.write "l_producto= "  & l_producto & "<br>"
'response.write "l_suma= "  & l_suma & "<br>"
'response.write "l_porcentajed= "  & l_porcentajed & "<br>"

%>

<table width="50" align="center">

<%
  
' POTENCIAL  --------------------------------------------------------
' primero busco si se definio una competencia potencial
  l_sql = " SELECT evatrvalor "
  l_sql = l_sql & " FROM  evaresultado "
  l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
  l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
  l_sql = l_sql & " INNER JOIN evafactor	 ON evaseccfactor.evafacnro = evafactor.evafacnro "
  l_sql = l_sql & "		 AND   evafactor.evafacpot = -1 " 
  l_sql = l_sql & " INNER JOIN evadetevldor	 ON evadetevldor.evldrnro = evaresultado.evldrnro "
  l_sql = l_sql & "		 AND   evadetevldor.evacabnro =  " &l_evacabnro
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =  " & l_evaseccnro 
'  Response.Write l_sql
  rsOpen l_rsx, cn, l_sql, 0
  if not l_rsx.eof then 
	 l_sumap = cdbl(l_rsx("evatrvalor")) * l_cantidad
  else	
	  l_rsx.close
	 'Si no hay copetencia potencial busco evaluador potencial
	  l_sql = " SELECT SUM(evatrvalor) AS suma "
	  l_sql = l_sql & " FROM  evaresultado "
	  l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
	  l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
	  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =  " & l_evaseccnro 
	  l_sql = l_sql & " AND   evaresultado.evldrnro = " & l_potencial
	  rsOpen l_rsx, cn, l_sql, 0
	  if not l_rsx.eof then 
		l_sumap = l_rsx("suma")
	  else
		Response.Write("<script>alert('Se necesitan valores Potenciales para realizar el gráfico.')</script>")
		Response.End
	  end if
	  l_rs.Close
  end if
  	  
if trim(l_sumap)="" or isnull(l_sumap)then
  Response.Write("<script>alert('Se necesitan valores Potenciales para realizar el gráfico.')</script>")
  Response.End
end if

'response.write "sumap= " & l_sumap & "<br>"

  if l_producto <> 0 then
	l_porcentajep   = cdbl(l_sumap)  / cdbl(l_producto)  * 100
  else 
	l_porcentajep   = 0
  end if

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
			if (cdbl(l_porcentajed) >= (100-i)) and (cdbl(l_porcentajed) > (100-i-5)) and ((cdbl(l_porcentajed)-100 + i ) < 5) then
				if  cdbl(l_porcentajep) >= j and (cdbl(l_porcentajep) < (j+5)) then%>
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

<%
end function
%>