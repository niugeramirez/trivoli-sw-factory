// Funcion para abrir ventanas centradas

//window.open(donde,"new019","toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=640,height=270");
function abrirVentana(url, name, width, height,opc) 
{
  var str = "height=" + height + ",innerHeight=" + height;
  str += ",width=" + width + ",innerWidth=" + width;
  if (window.screen) {
    var ah = screen.availHeight - 30;
    var aw = screen.availWidth - 10;

    var xc = (aw - width) / 2;
    var yc = (ah - height) / 2;
    if (xc < 0) 
		xc = 0
    if (yc < 0) 
		yc = 0
	str += ",left=" + xc + ",screenX=" + xc;
    str += ",top=" + yc + ",screenY=" + yc;
  }
//	alert(aw + ' ' + (aw - width) + ' ' + width) 
	str += ",resizable=yes"
	if (opc != null)
	  str += opc
    var auxi;
	auxi = url.substr(url.lastIndexOf('/')+1,url.length);
	auxi = auxi.substr(0,auxi.indexOf(".asp"));
	if (name!=""){
		auxi = name;
	}
	window.open(url, auxi, str);
//  window.open(url, url.substr((url.lastIndexOf("/")+1),url.indexOf(".asp")), str);
//  window.open(url, 'pepe', str);
}

function abrirVentanaCent(url, name, width, height,opc) 
{
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
//	alert(aw + ' ' + (aw - width) + ' ' + width) 
	str += ",resizable=yes";
	if (opc != null)
	   str += opc;
    var auxi;
	window.open(url, name, str);
//  window.open(url, url.substr((url.lastIndexOf("/")+1),url.indexOf(".asp")), str);
//  window.open(url, 'pepe', str);
}



function abrirVentanaVerif(url, name, width, height, opc, sistema) 
{
  if (ifrm.jsSelRow == null)
    alert("Debe seleccionar un registro.")
  else	
    if ((sistema != null) &&
        (ifrm.jsSelRow.cells(sistema).innerText.toUpperCase() == 'SI'))
       alert('Registro del sistema. No se lo puede modificar.');
	else   
       abrirVentana(url, url.substr(0,url.indexOf(".asp")), width, height, opc) 
}

//window.open(donde,"","toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=250,height=120");
function abrirVentanaH(url, name, width, height) 
{
  var str = "height=" + height + ",innerHeight=" + height;
  str += ",width=" + width + ",innerWidth=" + width;
  if (window.screen) {

    var xc = screen.availWidth + width;
    var yc = screen.availHeight + height;

    str += ",left=" + xc + ",screenX=" + xc;
    str += ",top=" + yc + ",screenY=" + yc;
  }
  window.open(url, name, str);
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
  return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString());
}

function Nuevo_DialogoH(w_in, pagina, ancho, alto)
{
  return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';dialogTop:300000;dialogLeft:300000');
}

//Funcion para cambiar el tamaño de una Ventana y Centrarla en la pantalla
function CentrarVentana(cx,cy){
	var ax;
	var ay;
	ax = (screen.availWidth - cx)/2;
	ay = (screen.availHeight - cy) /2;
	window.resizeTo(cx , cy);
	window.moveTo(ax,ay);
}


/*
==================================================================
LTrim(string) : Returns a copy of a string without leading spaces.
==================================================================
*/
function LTrim(str)
/*
   PURPOSE: Remove leading blanks from our string.
   IN: str - the string we want to LTrim
*/
{
   var whitespace = new String(" \t\n\r");

   var s = new String(str);

   if (whitespace.indexOf(s.charAt(0)) != -1) {
      // We have a string with leading blank(s)...

      var j=0, i = s.length;

      // Iterate from the far left of string until we
      // don't have any more whitespace...
      while (j < i && whitespace.indexOf(s.charAt(j)) != -1)
         j++;

      // Get the substring from the first non-whitespace
      // character to the end of the string...
      s = s.substring(j, i);
   }
   return s;
}

/*
==================================================================
RTrim(string) : Returns a copy of a string without trailing spaces.
==================================================================
*/
function RTrim(str)
/*
   PURPOSE: Remove trailing blanks from our string.
   IN: str - the string we want to RTrim

*/
{
   // We don't want to trip JUST spaces, but also tabs,
   // line feeds, etc.  Add anything else you want to
   // "trim" here in Whitespace
   var whitespace = new String(" \t\n\r");

   var s = new String(str);

   if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
      // We have a string with trailing blank(s)...

      var i = s.length - 1;       // Get length of string

      // Iterate from the far right of string until we
      // don't have any more whitespace...
      while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)
         i--;


      // Get the substring from the front of the string to
      // where the last non-whitespace character is...
      s = s.substring(0, i+1);
   }

   return s;
}

/*
=============================================================
Trim(string) : Returns a copy of a string without leading or trailing spaces
=============================================================
*/
function Trim(str)
/*
   PURPOSE: Remove trailing and leading blanks from our string.
   IN: str - the string we want to Trim

   RETVAL: A Trimmed string!
*/
{
   return RTrim(LTrim(str));
}
