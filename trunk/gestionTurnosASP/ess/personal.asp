<html>
<head>
<link href="<%= session("estilo")%>" rel="StyleSheet" type="text/css">
<title>Untitled Document</title>
</head>
<script>
	function agregar(){
		document.all.altas_02.src = 'personal_01.asp';
	}
</script>
<body>
<table cellpadding="0" cellspacing="0"  width="100%" >
	<tr>
		<td >
			<iframe src="personal_01.asp" name="abmpersonal" frameborder="0" scrolling="no"></iframe>
		</td>
	</tr>
	<tr>
		<td>
			<table width="100%" >
				<tr>
					<th>Altas y Bajas <a href="agregar();" target="altas_02">Agregar</a><a href="#">Modificar</a><a href="#">Eliminar</a></th>
				</tr>
				<tr>
					<td  width="100%">
						<!-- <iframe name="altas_02" frameborder="0" scrolling="no"></iframe> -->
						<iframe src="altas_00.asp" name="altas" frameborder="0" scrolling="no"></iframe>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</body>
<script>
		parent.document.all.principal.style.height = document.body.scrollHeight;
</script>
</html>
