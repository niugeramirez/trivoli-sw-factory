<% Option Explicit %>

<%
'Response.Redirect "http://bb-omh-pc009/oracle/"
'response.end
%> 

<html>
<head>
<link href="/intranet/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Intranet - Oleaginosa Moreno S.A.</title>
<script src="/intranet/shared/js/fn_windows.js"></script>
<script src="/intranet/shared/js/fn_confirm.js"></script>
<script src="/intranet/shared/js/fn_ayuda.js"></script>
<script>
function inicio(){
	var opc;
	var str;
	var height = 560;  //360 
	var width = 780;   //500
	var url = 'lanzador/lanzador2.asp';
//  var url = 'lanzador/lanzador2M.asp';  // Manera - MAmerica.
//	var url = 'lanzador.html';            // Resto de las plantas.
	var name = 'lanzador';
	var str = "height=" + height + ",innerHeight=" + height;
		str += ",width=" + width + ",innerWidth=" + width;
	if (window.screen) {
		var ah = screen.availHeight - 30;
		var aw = screen.availWidth - 10;
	
	    var xc = (aw - width) / 2;
	    var yc = (ah - height) / 2;
	    if (xc < 0) 
			xc = 0;
	    if (yc < 0) 
			yc = 0;
		str += ",left=" + xc + ",screenX=" + xc;
	    str += ",top=" + yc + ",screenY=" + yc;
	}
	str += ",resizable=no";
	if (opc != null)
	   str += opc;
	var auxi;
	window.open(url, name, str);
	window.opener = 'm';
	window.self.close()
}
</script>
</head>
<body onload="inicio();">
</body>
</html>
