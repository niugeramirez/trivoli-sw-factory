<% Option Explicit %>

<% 
 Dim l_wiznro
 Dim l_sql
 Dim l_rs
 Dim l_wizdesabr
 Dim l_programa


l_wizdesabr = "blabla"
l_programa = "portada_00.html" 

%>
<html>
<head>
<link href="../css/tablesraul.css" rel="StyleSheet" type="text/css">
<title>Titulo</title>
<script>
function ActPasos(codigo, clabel, nombre){
	//alert(codigo);
	document.pasos.location = "asisp_01.asp?wiznro=<%=l_wiznro%>&codigo="+codigo+"&label="+clabel+"&nombre="+nombre;
}

function RefrescarPasos() {
    //alert("Refrescandoup!!!");
    document.pasos.location.reload();
}

function Abrir(pagina, codigo, pasonro) {

  if (pagina.indexOf('?') < 0){
     document.ifrm.location = pagina+"?codigo="+codigo+"&pasnro="+pasonro;
  }else{
    document.ifrm.location = pagina+"&codigo="+codigo+"&pasnro="+pasonro;
  }    
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<table border="10" cellpadding="0" cellspacing="0" height="100%" width="100%">

	<tr>
    	<td align="center" class="barra" colspan="1" height="0" style="Navy; border-bottom: 1px solid White;"  nowrap>
		    <img style="filter:Shadow(Color=White,Direction=120);" src="../images/gen_rep/tablero.gif">
			
			<!--<SPAN 
			    STYLE="position: absolute; top:12px; left:300px; font-size: 18px; color: White; font-family: Arial, Helvetica, sans-serif;">
				   Asistente de <%'=l_wizdesabr%></SPAN> -->
				   
		</td>
		
	</tr>
<form name="datos" method="post">
<input type="hidden" name="pasonro" value="">
<input type="hidden" name="menunro" value="">
<input type="hidden" name="menunroant" value="">
</form>			
	<tr>
    	<td nowrap align="right" width="100%" height="70%" >&nbsp;
		<!--
			<iframe  name="ifrm" src="<%'=l_programa%>" width="100%" height="100%" frameborder="9" marginheight="50" marginwidth="50" scrolling="no"></iframe>
			-->
		</td>
	</tr>
	
	<tr>
		<td nowrap align="center" height="20%" width="100%">
			<iframe name="pasos" src="#" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" align="left"></iframe> 
		</td>
	</tr>
			

</table>
</body>
</html>
