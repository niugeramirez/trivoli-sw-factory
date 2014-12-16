<% Option Explicit %>
<!--
-----------------------------------------------------------------------------
Archivo        : gen_select_emp_04.asp
Creador        : Scarpa D.
Fecha Creacion : 25/11/2004
Descripcion    : Actualiza la lista de empleados
Modificacion   :
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_lista
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

Dim l_accion
Dim l_listaN
Dim l_empleg
Dim l_terAnt
Dim l_agregados

l_pagina      = CInt(request("paginader"))
l_vent_pagina = CInt(request("ventpagina"))
l_tipovent    = request("tipovent")

l_selalto    = request("selalto")
l_selancho   = request("selancho")

l_filtro = request("sqlfiltroder")
l_orden  = request("sqlordender")
l_lista  = request("seleccion")

l_accion = request("accion")
l_listaN = request("listanueva")

Dim l_arr
Dim l_i
Dim l_arrL
Dim l_arrN
Dim l_idx1
Dim l_idx2
Dim l_salir
Dim l_tmp1
Dim l_tmp2
Dim l_legajo1
Dim l_legajo2

'response.write "Nueva  : " & l_listaN & "<br>"
'response.write "Actual : " & l_lista & "<br>"
'response.write l_accion & "<br>"

l_arr = split(l_listaN,",")

select case l_accion
   'Sacar empleados de la seleccion
   case "Q" 
	    l_lista = l_lista & ",0"
		for l_i = 0 to Ubound(l_arr)
	        l_lista = replace(l_lista,"," & l_arr(l_i) & ",",",")
		next
		
	    l_lista = mid(l_lista,1,len(l_lista)-2)
		
	    if mid(l_lista,len(l_lista),1) = "," then
	       l_lista = mid(l_lista,1,len(l_lista)-1)
		end if

   'Agregar empleados a la seleccion
   case "A"
	    if len(l_lista) > 1 then
		
		   l_arrL = split(l_lista,",")
		   l_arrN = split(l_listaN,",")

		   l_lista = "0"

	       l_idx1 = 1
		   l_idx2 = 0
		   
		   l_salir = false
		   
		   l_i = 0

	       do until (l_idx1 > UBound(l_arrL)) OR (l_idx2 > UBound(l_arrN))
		   
		      l_tmp1 = split(l_arrL(l_idx1),"@")
			  l_tmp2 = split(l_arrN(l_idx2),"@")

		      l_legajo1 = CLng(l_tmp1(1))
			  l_legajo2 = CLng(l_tmp2(1))
			  
			  if l_legajo1 < l_legajo2 then
			     l_lista = l_lista & "," & l_arrL(l_idx1)
			     l_idx1 = l_idx1 + 1
			  else
			     l_lista = l_lista & "," & l_arrN(l_idx2)
			     l_idx2 = l_idx2 + 1
			  end if

		   loop
		   
		   if (l_idx1 > UBound(l_arrL)) then
		       do until (l_idx2 > UBound(l_arrN))
				     l_lista = l_lista & "," & l_arrN(l_idx2)
				     l_idx2 = l_idx2 + 1
			   loop
		   else
		       do until (l_idx1 > UBound(l_arrL))
				     l_lista = l_lista & "," & l_arrL(l_idx1)
				     l_idx1 = l_idx1 + 1
			   loop
		   end if
	
		else
		   l_lista = l_lista & "," & l_listaN 
		end if

end select

'response.write "Salida: " & l_lista
'response.end

%>  

<script>
parent.document.datos.seleccion.value = '<%= l_lista%>';
parent.actualizarSelectados();
if ( (parent.document.ifrmfiltros.datos.listanro.value != "") || (parent.document.all.sqlfiltrofijo.value != "")){
    parent.actualizar();
}
</script>
