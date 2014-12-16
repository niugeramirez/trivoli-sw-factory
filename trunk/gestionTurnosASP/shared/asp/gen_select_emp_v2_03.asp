<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : gen_select_emp_03.asp
Creador        : Scarpa D.
Fecha Creacion : 27/11/2003
Descripcion    : Lista de empleados derecha - empleados selectados
Modificacion   :
  23/12/2003 - Scarpa D. - Cambio en la forma de la pagina.
-----------------------------------------------------------------------------
-->
<% 

on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_lista
Dim l_listaN
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_canttotal
Dim l_cantFiltro

Dim l_selalto
Dim l_selancho

Dim l_pagina
Dim l_vent_pagina
Dim l_actual
Dim l_hay_datos
Dim l_tipovent

Dim l_arr
Dim l_arrTemp
Dim l_maxPagina

l_pagina      = CInt(request("paginader"))
l_vent_pagina = CInt(request("ventpagina"))
l_tipovent    = request("tipovent")

l_selalto    = request("selalto")
l_selancho   = request("selancho")

l_filtro = request("sqlfiltroder")
l_orden  = request("sqlordender")
l_lista  = request("seleccion")

if l_orden = "" then
   l_orden = " ORDER BY empleg"  'orden por defecto legajo
end if

if l_tipovent = "1" then
   l_vent_pagina = 30000
end if

Function URLDecode(strConvert)

Dim arySplit
Dim strHex
Dim strOutput
Dim i
Dim Letter

If IsNull(strConvert) Then
   URLDecode = ""
   Exit Function
End If

' First convert the + to a space
strOutput = REPLACE(strConvert, "+", " ")

' Then convert the %number to normal code
arySplit = Split(strOutput, "%")

If IsArray(arySplit) Then
   strOutput = arySplit(i)
   For I = LBound(arySplit) to UBound(arySplit) - 1
      strHex = "&H" & Left(arySplit(i+1),2)
      Letter = Chr(strHex)

      strOutput = strOutput & Letter & Right(arySplit(i+1),len(arySplit(i+1))-2)
   Next
End If

URLDecode = strOutput

End Function

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<html>
<head>
	<title><%= Session("Titulo")%>Untitled</title>
<script languaje="javascript">
function Cargar1(){
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_lista  = mid(l_lista,3,len(l_lista)-1)

l_arr = split(l_lista,",")

if UBound(l_arr) = -1 then

    l_canttotal = 0

else
	'Controlo que la pagina actual no sea mayor a la ultima
	if ((UBound(l_arr) + 1) mod l_vent_pagina) = 0 then
	   l_maxPagina = Int(((UBound(l_arr) + 1) / l_vent_pagina))
	else 
	   l_maxPagina = Int(((UBound(l_arr) + 1) / l_vent_pagina) + 1) 
	end if
	
	if CInt(l_pagina) > CInt(l_maxPagina) AND l_maxPagina <> 0 then
	   l_pagina = l_maxPagina
	end if

	l_canttotal = 0
	'Salteo los registros hasta llegar a la pagina indicada
	if CInt(l_pagina) = 1 then
	   l_canttotal = 0
	else
	   l_canttotal = (l_vent_pagina * (l_pagina - 1)) 
	end if
	
	l_actual = 0
	l_hay_datos = false
	
	l_cantfiltro = -1
	
	if l_canttotal <= UBound(l_arr) then
	
	    l_listaN = ""
	
		do until l_actual = l_vent_pagina
			
		    l_arrTemp = split(l_arr(l_canttotal),"@")
			
			if l_listaN = "" then
			   l_listaN = l_arrTemp(0)
			else
			   l_listaN = l_listaN & "," & l_arrTemp(0)
			end if
			
		    l_canttotal = l_canttotal + 1
			
			if  l_canttotal > UBound(l_arr)  then
			   exit do
			else
			   l_actual = l_actual + 1 
			end if
		loop
		
		l_hay_datos = (l_actual = l_vent_pagina)
		
		'Cuento la cantidad de registros que quedan
		l_canttotal = UBound(l_arr) + 1
	
		'Busco los datos de los empleados que entran en la ventana
		
		l_sql = "SELECT DISTINCT ternro,empleg, terape, ternom "
		l_sql = l_sql & "FROM v_empleado "
		l_sql = l_sql & "WHERE ternro IN (" & l_listaN & ") "
		
		if l_filtro <> "" then
		  l_sql = l_sql & " AND " &  URLDecode(l_filtro) & " "
		end if
		
		l_sql = l_sql & " " & l_orden
		
		rsOpen l_rs, cn, l_sql, 0 
		
		do until l_rs.eof
		
		    response.write "newOp = new Option();" & vbCrLf
		    response.write "newOp.value  = '" & l_rs("ternro") & "@" & l_rs("empleg") & "';" & vbCrLf
		    response.write "newOp.text   = '" & l_rs("empleg") & " - " & l_rs("terape") & ", " & l_rs("ternom") & "';"  & vbCrLf
		    l_cantfiltro = l_cantfiltro + 1
		    response.write "document.registro.selfil.options[" & l_cantfiltro & "] = newOp;" & vbCrLf
		
			l_rs.MoveNext
		loop
		
		l_rs.close
	end if
	
	set l_rs = Nothing
	cn.Close
	set cn = Nothing

end if
%>  
}

</script>	
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" scroll=no>
<form name="registro">
<input type="Hidden" name="lista" value="<%= l_lista %>">
<select multiple size="17" style="width:<%= l_selancho%>px;height:<%= l_selalto%>px" name="selfil" ondblclick="parent.Uno(selfil,parent.nselfil.registro.nselfil, parent.document.datos.totalder, parent.document.datos.totalizq);"></select>
</form>
<script>

Cargar1();

<%if CInt(l_cantfiltro) = -1  then%>
  parent.document.datos.filtroder.value = 'No hay datos.';
<%else%>
  parent.document.datos.filtroder.value = '<%= CInt(l_cantfiltro) + 1 %>';
<%end if%>

<% if (l_canttotal / l_vent_pagina) = 0 then%>
   parent.document.datos.totpaginader.value = 1
<% else %>
	<% if (l_canttotal mod l_vent_pagina) = 0 then%>
	   parent.document.datos.totpaginader.value = <%= Int(l_canttotal / l_vent_pagina)  %>
	<% else %>
	   parent.document.datos.totpaginader.value = <%= Int(l_canttotal / l_vent_pagina) + 1 %>
	<% end if %>
<% end if %>
parent.document.datos.paginader.value = '<%= l_pagina%>';
<%'if l_filtro = "" then%>
parent.document.datos.totalder.value = '<%= l_canttotal %>';	
<%'end if%>
<%if l_hay_datos then%>
parent.document.datos.paginaDerFin.value = "0";
<%else%>
parent.document.datos.paginaDerFin.value = "1";	
<%end if%>

</script>
</body>
</html>
