<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/sqls.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : gen_select_emp_02.asp
Creador        : Scarpa D.
Fecha Creacion : 27/11/2003
Descripcion    : Lista de empleados izquierda - empleados no selectados
Modificacion   :
-----------------------------------------------------------------------------
-->
<% 

on error goto 0

Dim l_rs
Dim l_sql

Dim l_filtro
Dim l_lista
Dim l_lista2
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_canttotal
Dim l_cantFiltro
Dim l_hay_datos
Dim l_tipovent
Dim l_maxPagina

dim l_posicion
dim l_auxiliar

Dim l_seleccion
Dim l_sqlfiltroemp
Dim l_sqlfiltrofijo
Dim l_sqloperando
Dim l_hay_filtro

Dim l_selalto
Dim l_selancho

Dim l_pagina
Dim l_vent_pagina
Dim l_actual

l_pagina      = CInt(request("paginaizq"))
l_vent_pagina = CInt(request("ventpagina"))
l_tipovent    = request("tipovent")

l_selalto    = request("selalto")
l_selancho   = request("selancho")

l_seleccion     = request("seleccion")
l_sqlfiltroemp  = request("sqlfiltroemp")
l_sqlfiltrofijo = request("sqlfiltrofijo")
l_sqloperando   = request("sqloperando")

l_filtro = request("sqlfiltroizq")
l_orden  = request("sqlordenizq")

l_orden = " ORDER BY empleg"

if l_tipovent = "1" then
   l_vent_pagina = 30000
end if

'-----------------------------------------------------------------------------------
' Descripcion: Cambia los caracteres puestos por Escape
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


Dim l_idx_dest
Dim l_arrDest1(10000)
Dim l_arrDest2(10000)

'-----------------------------------------------------------------------------------
' Descripcion: Saca de los empleados encontrados los que ya estan seleccionados
sub generarArreglo()
    Dim l_avanzar
	Dim l_idx_sel
	Dim l_idx_fil
	Dim l_arrSel
	Dim l_arrTemp

    l_idx_sel  = 1
    l_idx_dest = 0

	if l_seleccion = "" then
       l_seleccion = "0@0" 
	   l_arrSel = split("0@0",",")
	else
	   l_arrSel = split(l_seleccion,",")
	end if

    do until l_rs.eof
       l_avanzar = false

	   if l_idx_sel > UBound(l_arrSel) then
	      l_avanzar = true
	   else
 	      l_arrTemp = split(l_arrSel(l_idx_sel),"@")
		  
	      if CLng(l_rs("ternro")) <> CLng(l_arrTemp(0)) then
		     if CLng(l_rs("empleg")) < CLng(l_arrTemp(1)) then
                l_avanzar = true
			 else
		        l_idx_sel = l_idx_sel + 1
			 end if
		  else
		     l_idx_sel = l_idx_sel + 1
			 l_rs.movenext
		  end if
	   end if

	   if l_avanzar then
'	   			response.write l_idx_dest
'				response.write l_rs("empleg")
'			response.end

	      l_arrDest1(l_idx_dest) = l_rs("ternro") & "@" & l_rs("empleg")
		  l_arrDest2(l_idx_dest) = l_rs("empleg") & " - " & l_rs("terape") & ", " & l_rs("ternom")
		  l_idx_dest = l_idx_dest + 1
		  l_rs.movenext
	   end if

    loop

    l_rs.close

end sub 'generarArreglo()

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

'response.write "parent.totalizq.value = " & rs(0) & ";" & vbCrLf

' Armo la consulta SQL
l_sql =         " SELECT DISTINCT * "
l_sql = l_sql & " FROM v_empleado "
l_sql = l_sql & " WHERE "

Dim l_arr
Dim l_i

l_arr = split(l_sqlfiltroemp,";")

l_sql = l_sql & " ( "

l_hay_filtro = false

if l_sqlfiltrofijo <> "" then
   l_hay_filtro = true
   l_sql = l_sql & " ternro IN ( "
   l_sql = l_sql & l_sqlfiltrofijo
   l_sql = l_sql & " ) "
end if

' Busco todos los filtros
for l_i = 0 to UBound(l_arr)
    if l_hay_filtro then
       l_sql = l_sql & l_sqloperando & " ternro IN ( "
       l_sql = l_sql & l_arr(l_i) 
       l_sql = l_sql & " ) "
	else
	   l_hay_filtro = true
       l_sql = l_sql & " ternro IN ( "
       l_sql = l_sql & l_arr(l_i) 
       l_sql = l_sql & " ) "
	end if
next

l_sql = l_sql & " ) "

' Elimino de la seleccion los elementos seleccionados
'if l_seleccion <> "" then
'   l_sql = l_sql & " AND ternro NOT IN ( "
'   l_sql = l_sql & l_seleccion
'   l_sql = l_sql & " ) "
'end if

l_hay_datos = false
l_canttotal = 0

if l_hay_filtro then

    if l_filtro <> "" then
       l_sql = l_sql & "AND " & URLDecode(l_filtro) & " "
    end if

    l_sql = l_sql & l_orden
	
	rsOpen l_rs, cn, l_sql, 0 
	
	if not l_rs.eof then
	
		'Genero el arreglo con los empleados del filtro, sacando los que estan seleccionados
		generarArreglo()
		
		l_cantfiltro = -1
		
		'Controlo que la pagina actual no sea mayor a la ultima
		if (l_idx_dest mod l_vent_pagina) = 0 then
		   l_maxPagina = Int((l_idx_dest / l_vent_pagina) )
		else 
		   l_maxPagina = Int((l_idx_dest / l_vent_pagina) + 1 )
		end if
		
		if CInt(l_pagina) > CInt(l_maxPagina) AND l_maxPagina <> 0 then
		   l_pagina = l_maxPagina
		end if

	    'Salteo los registros hasta llegar a la pagina indicada
		if CInt(l_pagina) = 1 then
	 	   l_canttotal = 0
		else
		   l_canttotal = (l_vent_pagina * (l_pagina - 1)) 
		end if
		
'	response.write l_idx_dest & "<br>"		
'	response.write l_canttotal & "<br>"			
'	response.write l_seleccion & "<br>"	
'	response.write UBound(l_arrSel) & "<br>"	
'	response.end
		
	
		l_actual = 0
		if l_idx_dest > l_canttotal then
		
			do until l_actual = l_vent_pagina
			
			    response.write "newOp = new Option();" & vbCrLf
			    response.write "newOp.value  = '" & l_arrDest1(l_canttotal) & "';" & vbCrLf
			    response.write "newOp.text   = '" & l_arrDest2(l_canttotal) & "';"  & vbCrLf
			    l_cantfiltro = l_cantfiltro + 1
			    response.write "document.registro.nselfil.options[" & l_cantfiltro & "] = newOp;" & vbCrLf
				
			    l_canttotal = l_canttotal + 1
				
				if l_idx_dest <= l_canttotal then
				   exit do
				else
				   l_actual = l_actual + 1 
				end if
			loop
			
			l_hay_datos = (l_actual = l_vent_pagina)
			
			'Cuento la cantidad de registros que quedan
			l_canttotal = l_idx_dest
'			response.write l_idx_dest
'			response.end
			
		end if

	end if
end if
	
set l_rs = Nothing
cn.Close
set cn = Nothing

if l_idx_dest = "" then
   l_idx_dest = 0
end if

%>  
}

</script>	
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" scroll=no>

<form name="registro">
<select multiple size="17" style="width:<%= l_selancho%>px;height:<%= l_selalto%>px" name=nselfil ondblclick="parent.Uno(nselfil,parent.selfil.registro.selfil, parent.document.datos.totalizq, parent.document.datos.totalder);"></select>
</form>
<script>

	Cargar1();

	<%if CInt(l_cantfiltro) = -1  then%>
	  parent.document.datos.filtroizq.value = 'No hay datos.';
	<%else%>
	  parent.document.datos.filtroizq.value = '<%= CInt(l_cantfiltro) + 1 %>';
	<%end if%>	
    
	<% if (l_canttotal / l_vent_pagina) = 0 then%>
	   parent.document.datos.totpaginaizq.value = 1
	<% else %>
		<% if (l_canttotal mod l_vent_pagina) = 0 then%>
		   parent.document.datos.totpaginaizq.value = <%= Int(l_canttotal / l_vent_pagina)  %>	
		<% else %>
		   parent.document.datos.totpaginaizq.value = <%= Int(l_canttotal / l_vent_pagina) + 1 %>	
		<% end if %>
	<% end if %>
    parent.document.datos.paginaizq.value = <%= l_pagina%>;
		
	<%'if l_filtro = "" then%>
	parent.document.datos.totalizq.value = '<%= l_idx_dest%>';	
	<%'end if%>
	<%if l_hay_datos then%>
	parent.document.datos.paginaIzqFin.value = "0";
	<%else%>
	parent.document.datos.paginaIzqFin.value = "1";	
	<%end if%>
</script>
</body>
</html>
