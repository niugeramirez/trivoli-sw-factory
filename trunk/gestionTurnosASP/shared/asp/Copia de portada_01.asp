<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
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
Dim l_mes
Dim l_cant_registros

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>s</title>

<script language="javascript">   
  
function repetirError() {  
    setTimeout("efectoImagen()",2000);   
    return true;   
}  
  
window.onerror = repetirError   
  
</script>  
  
<script language="javascript">   
  
// Creamos las imágenes que utilizaremos 
  
var image1 = new Image()  

// Definimos la URL de la imagen   

image1.src = "0.jpg"  

// Definimos el enlace que se cargará al hacer click sobre la imagen   

link1 = 'http://enlace1.html'  

var image2 = new Image()  

image2.src = "1.jpg"  

link2 = 'http://enlace2.html'  

var image3 = new Image()  

image3.src = "2.jpg"  

link3 = 'http://enlace3.html'  
  
</script> 

</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
 fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
    if (jsSelRow != null) {
        Deseleccionar(jsSelRow);
    };
 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
 <% 'If l_asistente = 1 then %>
    //parent.parent.ActPasos(cabnro,"","buques");
    //parent.parent.datos.pasonro.value = cabnro;
 <%' End If %>
}

function posY(obj){
  return( obj.offsetParent==null ? obj.offsetTop : obj.offsetTop+posY(obj.offsetParent) );
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" bgcolor="#000000" onLoad="efectoImagen()">
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="desc" value="">
<input type="hidden" name="orden" size="50" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
<!--
<table>
    <tr>
        <th width="100%" colspan="4">Cumpleaños del Mes</th>		
    </tr>
	
    <tr>
        <th width="40%">Apellido</th>
        <th width="40%">Nombre</th>
        <th width="20%">Fec. Nac.</th>		
    </tr>
-->

<div align="center">   
<A href="javascript:efectoLink()" name="enlace">   
<img src="0.jpg" name="efecto" border="0" style="filter:blendTrans(duration=3); ">   
</A>   
</div>   

 
<script>    
	//alert('kiki');
	parent.parent.ActPasos(1,"","");
	//parent.parent.datos.pasonro.value = <%'= l_primero %>;
</script>
<!--
</table>
-->


<script language="javascript">   
  
// Definimos si queremos utilizar la imagen como enlace hacia otra página   
var enlace = true // Valores: true || false   
// Cantidad de imágenes a utilizar en el efecto   
var numImagenes = 3   
// Velocidad del efecto en segundos   
var velocidad = 3   
  
// Definimos que imagen se carga primero   
var pasoImagen = 2   
  
function efectoImagen() {  
  
    if (!document.images) {  
        return }  
    if (document.all) {  
        efecto.filters.blendTrans.apply()  
        document.images.efecto.src = eval("image"+pasoImagen+".src") }  
    if (document.all) {  
        efecto.filters.blendTrans.play() }  
    if (pasoImagen < numImagenes) {  
        pasoImagen++ }  
    else {  
        pasoImagen = 1 }  
           
    if (document.all) {  
        setTimeout("efectoImagen()",velocidad*1000+3000) }  
    else {  
        setTimeout("efectoImagen()",velocidad*1000) }  
  
} // Fin de la función efectoImagen()   
  
function efectoLink() {  
  
if (enlace) {  
  
var imgCargada = document.images.efecto.src  
pasoImagenTemp = imgCargada.substring(imgCargada.length-5,imgCargada.length-4)  
  
    if (pasoImagenTemp == '0') {  
        window.location = link1 }  
    else if (pasoImagenTemp == '1') {  
        window.location = link2 }  
    else {  
        window.location = link3 }  
}  
  
else { return }  
  
} // Fin de la función efectoLink()   
  
</script>   

</body>
</html>
