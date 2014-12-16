<% Option Explicit %>

<% 
on error goto 0
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_asistente
Dim l_primero
Dim l_primerob
Dim l_primeroc
Dim l_codigo
Dim l_reemplazaestrnro


if l_orden = "" then
  l_orden = " ORDER BY int_sugerencia.sugfec desc "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_grayraul.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>buques - Oleaginosa Moreno Hnos. S.A.</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
 fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,desc, codext){
    if (jsSelRow != null) {
        Deseleccionar(jsSelRow);
    };
 document.datos.cabnro.value = cabnro;
 document.datos.desc.value = codext;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
 <% 'If l_asistente = 1 then %>
    //parent.parent.ActPasos(cabnro,"Puestos",desc);
    //parent.parent.datos.pasonro.value = cabnro;
 <% 'End If %>
}

function posY(obj){
  return( obj.offsetParent==null ? obj.offsetTop : obj.offsetTop+posY(obj.offsetParent) );
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" bgcolor="#000000">
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="desc" value="">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
<table border="0">
    <tr>
        <th nowrap width="100%" colspan="3" align="left">Listado de Sugerencias</th>
    </tr>
</table>
</body>
</html>
