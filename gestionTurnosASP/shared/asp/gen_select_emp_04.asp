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

l_arr = split(l_listaN,",")

if l_accion = "Q" then
    l_lista = l_lista & ",0"
	for l_i = 0 to Ubound(l_arr)
        l_lista = replace(l_lista,"," & l_arr(l_i) & ",",",")
	    response.write l_lista & "<br>"			
	next
	
    l_lista = mid(l_lista,1,len(l_lista)-2)
else
    if len(l_lista) > 1 then
       l_lista = "0," & l_listaN & "," & mid(l_lista,3,len(l_lista)-2)
	else
       l_lista = l_lista & "," & l_listaN 
	end if
end if


%>  

<script>
parent.document.datos.seleccion.value = '<%= l_lista%>';

parent.actualizarSelectados();

if (parent.document.nselfil.datos.selfil.length > 0){
    parent.actualizar();
}

</script>
