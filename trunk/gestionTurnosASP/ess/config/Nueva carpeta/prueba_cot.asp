<%@ LANGUAGE = VBScript %>
<%
	Option Explicit

	' Variables de aplicacion
	Dim strServer, strPathDsi, strAp

%>

<html>
<head>
	<title>Untitled</title>
</head>

<body>

<form action="http://cottest.ec.gba.gov.ar/TransporteBienes/SeguridadCliente/presentarRemitos.do" method="post" enctype="multipart/form-data">
<input type="hidden" name="user" value="33502232229">
<input type="hidden" name="password" value="502232">
<input type="hidden" name="file" value="c:\TB_30505241033_000000_20070103_000031.txt">
<p><input type="submit" value="Enviar"></p>
<INPUT type="button" value="Volver" tabIndex="11" onclick="javascript:document.location='<%=strPathDsi%>'" id=button1 name=button1></INPUT>										
</form>
</body>
</html>
