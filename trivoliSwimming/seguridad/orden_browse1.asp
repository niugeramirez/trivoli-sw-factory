<% Option Explicit %>
<%
Dim l_lista 
Dim l_campos 
Dim l_filtro 
Dim l_destino 
Dim l_cantidad
Dim l_actual
Dim l_listado(20)
Dim l_listaCampos(20)
Dim l_i

l_lista   = Request.querystring("Lista")
l_campos  = Request.querystring("Campos")
l_filtro  = Request.querystring("Filtro")
l_destino = Request.querystring("Pagina")

l_cantidad = 0

do while len(l_campos) > 0
  if inStr(l_campos,";") <> 0 then
    l_actual  = left(l_campos, inStr(l_campos,";") - 1)
    l_campos  = mid (l_campos, inStr(l_campos,";") + 1)
  else
    l_Actual = l_campos
	l_campos = ""
  end if
  l_cantidad = l_cantidad + 1
  l_listaCampos(l_cantidad) = l_actual
loop

l_cantidad = 0

do while len(l_lista) > 0
  if inStr(l_lista,";") <> 0 then
    l_actual = left(l_lista, inStr(l_lista,";") - 1)
    l_lista  = mid (l_lista, inStr(l_lista,";") + 1)
  else
    l_Actual = l_lista
	l_lista = ""
  end if
  l_cantidad = l_cantidad + 1
  l_listado(l_cantidad) = l_actual
loop

%>

<html>
<head>
<link href="/trivoliSwimming/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Orden - Ticket</title>
</head>
<script>
function valor()
{
  for (i=0;i<=<%= l_cantidad * 2 - 1 %>;i++)
  {
    if (document.datos.orden[i].checked)
	  var sel = document.datos.orden[i].value
  }
  if (sel.substr(0,1) == "-")
    {
	sel = sel.substr(1) + ' DESC'
	if (sel.search(',') >= 0)
	  {
	  sel = sel.substr(0,sel.search(",")) + " DESC" + sel.substr(sel.search(","))
	  } 
	}
  sel = "ORDER BY "	+ sel
  return sel
}
window.resizeTo(360,<%= l_cantidad * 25 + 100 %>)
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<form name="datos" method="post">
<input type="Hidden" name="filtro" value='<%= l_filtro %>'>
<table cellspacing="1" cellpadding="0" border="0" width="100%">
  <tr>
    <td class="th2" colspan="3">Orden</td>
  </tr>
  <tr>
    <td class="th2">&nbsp;</td>
    <td class="th2">Ascendente</td>
    <td class="th2">Descendente</td>
  </tr>
<%
l_i = 1
do while l_i <= l_cantidad
%>
<tr>
    <td align="right"><b><%= l_listado(l_i)%></b></td>
	<td><input type="Radio" name="orden" value="<%= l_listaCampos(l_i) %>" <% If l_i = 1 then%>checked<% End If %>></td>
	<td><input type="Radio" name="orden" value="-<%= l_listaCampos(l_i) %>"></td>
</tr>
<%
  l_i = l_i + 1
loop
%>  
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:window.opener.ifrm.location = '<%= l_destino %>&orden=' + valor() + '&filtro=' + document.datos.filtro.value; window.close();">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>		
	</td>
</tr>
</table>
</form>
</body>
</html>
