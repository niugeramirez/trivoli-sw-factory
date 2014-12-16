<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<html>
<head>
<link href="<%= c_estilo%>" rel="StyleSheet" type="text/css">
<title>Principal</title>
</head>
<script>
	function opciones(){
		document.all.principal.src='personal.asp';
	}
	function Click(obj){
		 for (var i=0; i<document.links.length; i++){
			document.links[i].className="boton";
		}
		//document.link.style.className = 'boton';
		obj.className = 'botonsel';
	}
	
	function adjustIFrameSize (iframeWindow) {
	  if (iframeWindow.document.height) {
	     var iframeElement = document.getElementById(iframeWindow.name);
	     iframeElement.style.height = iframeWindow.document.height + 'px';
	     iframeElement.style.width = iframeWindow.document.width + 'px';
	  }
	  else if (document.all) {
	    var iframeElement = document.all[iframeWindow.name];
	    if (iframeWindow.document.compatMode &&
	        iframeWindow.document.compatMode != 'BackCompat') 
	    {
	      iframeElement.style.height = iframeWindow.document.documentElement.scrollHeight + 5 + 'px';
	      iframeElement.style.width  = iframeWindow.document.documentElement.scrollWidth + 5 + 'px';
	    }
	    else {
	      iframeElement.style.height = iframeWindow.document.body.scrollHeight + 5 + 'px';
	      iframeElement.style.width = iframeWindow.document.body.scrollWidth + 5 + 'px';
	    }
	  }
	}	
	
</script>
<body class="indexprincipal" onload="if (parent.adjustIFrameSize) parent.adjustIFrameSize(window);">
	<table class="Tprincipal" cellpadding="0" cellspacing="0">
 		<tr>
			<td class="barmenu">&nbsp;
				 <!--<h3>Bienvenidos al servicio de AUTOGESTIÓN</h3> -->
			</td>
		</tr>
		<tr>
			<td class="tdprincipal">
				<iframe name="principal" id="principal" src="bienvenida.asp" frameborder="0" scrolling="yes"></iframe>
			</td>
		</tr>
	</table>
</body>
</html>

