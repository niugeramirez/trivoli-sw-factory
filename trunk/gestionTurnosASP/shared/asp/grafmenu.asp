<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!-------------------------------------------------------------------------------
Archivo		: grafmenu.asp
Descripción	: Barra de Menú para los gráficos de web components.
Autor		: Lisandro Moro
Fecha		: 13/09/2004
------------------------------------------------------------------------------->
<html>
<head>
	<title><%= Session("Titulo")%>Menú Gráficos</title>
</head>
<%
'on error goto 0
Dim l_estilo_a
Dim l_estilo_i
Dim l_class
Dim l_estilo	'Recive los parametro del estilo que se deben mostrar
Dim l_graficos	'Recive los parametro de los graficos que se deben mostrar

l_graficos = request.querystring("graficos")
l_estilo = request.querystring("estilo")

if l_estilo <> "" then
	l_estilo = l_estilo
else
	l_estilo = "tables4.css"
end if
%>
<link href="/turnos/shared/css/tables4.css" rel="StyleSheet" type="text/css" id="csss">
<style type="text/css">
img{border:none;}
td{}
a{
	background-image:url(../images/graf_0a.gif);
	height: 100%;
	padding : 0 0 0 0;
	width: 30px;
}
A:hover{
	background-image:url(../images/graf_0b.gif);
	height: 100%;
	padding : 0 0 0 0;
	width: 30px;
}
.anormal{background-image:url(../images/graf_0a.gif);}
.ainversa{background-image:url(../images/graf_0b.gif);}
</style>
<script>
var txt;
var j;
var count;

function actualiza(nro){
    for (j=0; j < document.links.length; j++) {
		if (document.links(j).id == nro){
			document.links(j).className = "ainversa";
		}else{
			document.links(j).className = "anormal";
		}
	}
	parent.generar(nro);
}
</script>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<% 
'--estilo para lo links--
l_estilo_a = "height:100%;width:30;border:none;vertical-align:middle;"
'--estilo para las imagenes--
l_estilo_i = "height:95%;width:25;vertical-align:middle;"
	
Dim l_arreglo
Dim l_arreglo_graficos
	
function linea(a)
	Response.write "<td align=""center"" valign=""middle""><a id=" & l_arreglo(a) & " class="& l_class &" style="& l_estilo_a & " href=""Javascript:actualiza("& l_arreglo(a) &")""><img src=""../images/graf_" & l_arreglo(a) & ".gif"" style="& l_estilo_i &"></a></td>" & vbcrlf
end function
	
function Cuales
	Dim a
	Dim b
	l_arreglo = Array(0,1,3,4,6,14,18,19,20,29,31,32,33,34,36)
	if l_graficos = "" then 'si no vienen parametros
		for a = 0 to UBound(l_arreglo)
			linea(a)
		next
	else 'si se paarametrizan los botones
		l_arreglo_graficos = split(l_graficos, ";")
		for a = 0 to UBound(l_arreglo_graficos)
			For b = 0 to UBound(l_arreglo)
				if CInt(l_arreglo(b)) = CInt(l_arreglo_graficos(a)) then
					linea(b)
				end if
			next
		next
	end if
end function	%>
<table style="height:25;" height="15" width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td width="50%">&nbsp;</td>
		<td style="background-image:url(../images/graf_0a.gif);">
			<table id="tabla" style="border: thin solid Silver;" cellpadding="0" cellspacing="0" border="0" height="30">
				<tr valign="middle">
					<% Cuales %>
				</tr>
			</table>
		</td>
		<td width="50%">&nbsp;</td>
	</tr>
</table>
</body>
</html>
