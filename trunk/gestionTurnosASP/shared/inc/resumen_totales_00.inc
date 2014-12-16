<%



Dim l_producto
Dim l_porcentaje
Dim l_puntajefinal
Dim l_suma
function calcular(evaseccnro,evaluador)
Dim l_cantidad
Dim l_maximo
l_cantidad= 0
l_suma=0
l_maximo=0

dim l_rs1


  Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
' Calcular Suma Total
  l_sql = " SELECT SUM(evatrvalor) AS suma "
  l_sql = l_sql & " FROM  evaresultado "
  l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
  l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =  " & evaseccnro 
  l_sql = l_sql & " AND evaresultado.evldrnro = " & evaluador

  rsOpen l_rs1, cn, l_sql, 0
  if not l_rs1.eof then 
	l_suma = l_rs1("suma")
  end if
  l_rs1.Close
  
' Buscar el maximo valor    
  l_sql = "SELECT MAX(evatrvalor) AS maximo "
  l_sql = l_sql & " FROM  evatipresu "
  l_sql = l_sql & " INNER JOIN evaresu ON evaresu.evatrnro = evatipresu.evatrnro "
  l_sql = l_sql & " WHERE evaresu.evaseccnro = " & evaseccnro

  rsOpen l_rs1, cn, l_sql, 0
  if not l_rs1.eof then 
	l_maximo = l_rs1("maximo")
  end if
  l_rs1.Close

' Buscar cantidad de competencias para la seccion  
  l_sql = "SELECT count(evafacnro) AS cantidad "
  l_sql = l_sql & " FROM  evaseccfactor "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & evaseccnro

  rsOpen l_rs1, cn, l_sql, 0
  if not l_rs1.eof then 
	l_cantidad = l_rs1("cantidad")
  end if
  l_rs1.Close

l_producto = cdbl(l_maximo) * cdbl(l_cantidad)


l_porcentaje = cdbl(l_suma)  / cdbl(l_producto)  * 100
 
l_puntajefinal = cdbl(l_suma) / cdbl(l_cantidad)
end function


function resumen_totales

DIM l_rsl
Set l_rsl = Server.CreateObject("ADODB.RecordSet")

l_sql= "select evasecc.evaseccnro, evadetevldor.evldrnro, evatipevalua.evatevdesabr "
l_sql= l_sql & "FROM  evadetevldor "
l_sql= l_sql & "INNER JOIN evasecc ON evadetevldor.evaseccnro=evasecc.evaseccnro "
l_sql= l_sql & "INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
l_sql= l_sql & "WHERE evadetevldor.evacabnro=" & l_evacabnro
l_sql= l_sql & " and evasecc.tipsecnro = 2 "

rsOpen l_rsl, cn, l_sql, 0 
%>
<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
<%
do until l_rsl.eof
	calcular l_rsl("evaseccnro"),l_rsl("evldrnro")
%>
	<tr>
		<td nowrap colspan="3"> <b><%= l_rsl("evatevdesabr") %></b> </td>
	</tr>
	<td align=center>Suma Total</td> 
    <td align=center>% Total</td>
	<td align=center>Calif. Final</td>
    <tr>
	<td nowrap align=center><%=l_suma%></td>
	<td nowrap align=center><%=FormatNumber(l_porcentaje,2)%>%</td>
	<td nowrap align=center><%=int(l_puntajefinal)%>&nbsp; (<%=FormatNumber(l_puntajefinal,2)%>)</td>
	</tr>
<%
	l_rsl.MoveNext
loop
%>
	</table>
<%

l_rs1.Close
end function

%>
