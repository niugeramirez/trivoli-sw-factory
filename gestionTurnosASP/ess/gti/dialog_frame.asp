<HTML>
<HEAD>
</HEAD>
<title>Ayuda - RHPro &reg;</title>
<%
Dim linea, pagina
pagina = request("pagina")
linea  = request.querystring
if linea <> "" then
  pagina = pagina + "?" + linea
end if  
%>
<frameset rows="100%,*">
    <frame name="2carac" src="<%=pagina%>" marginwidth="0" marginheight="0" scrolling="auto" frameborder="0">
	<noframes>
	<BODY>
	Error: Su explorador no admite frames.
	</BODY>
	</noframes>
</frameset>
</HTML>
