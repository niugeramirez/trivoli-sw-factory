<html>
<head>
<link href="shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title></title>
</head>
<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}
function Seleccionar(fila){
	if (jsSelRow != null) {
		Deseleccionar(jsSelRow);
	};
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
	<table width="100%" class="">
		<tr>
			<th>Alta</th>
			<th>Baja</th>
			<th>Causa</th>
			<th>Estado</th>
		</tr>
		<% dim a
		for a = 0 to 50 %>
		<tr onclick="Javascript:Seleccionar(this);">
			<td><%= "Alta " & a %></td>
			<td><%= "baja " & a %></td>
			<td><%= "estado " & a %></td>
			<td><%= now %></td>
		</tr>	
		<% next %>
	</table>
</body>
<script>
	parent.document.all.altas.style.height = document.body.scrollHeight;
	parent.parent.document.all.principal.style.height = document.body.scrollHeight;
	parent.parent.parent.document.all.centro.style.height = 1000;
	//parent.parent.parent.document.height = 160000;//document.body.scrollHeight;
</script>
</html>
