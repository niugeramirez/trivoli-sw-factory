/*  FUNCIONES:
	emailValido(valor)
	telefonoValido(arreglo)
	nombreValido(nombre)
	stringValido(nombre)
	ValidaCuit(nroCuit)
*/

function emailValido(valor) {
  if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(valor)){
    return (true)
  } else {
    return (false);
  }
}

function telefonoValido(arreglo){
var n = arreglo.length
var i
//Los caracteres válidos son números,(,),-,*,# y espacio
for (i=0;i<n;i++)
	if (!((arreglo.charCodeAt(i)>47 && arreglo.charCodeAt(i)<58)||(arreglo.charCodeAt(i)>39 && arreglo.charCodeAt(i)<43)||(arreglo.charCodeAt(i)==45)||(arreglo.charCodeAt(i)==32)||(arreglo.charCodeAt(i)==35))){
		return(false)
	}
return(true)
}

function nombreValido(nombre){
  var checkOK = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZÄÁËÉÏÍÖÓÜÚ" + "abcdefghijklmnñopqrstuvwxyzäáëéïíöóüú ()*´`";
  var valido = true; 
  for (i = 0; i < nombre.length; i++) {
    ch = nombre.charAt(i); 
    for (j = 0; j < checkOK.length; j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length) { 
      valido = false; 
      break; 
    }
  }
  if (!valido) 
	return (false); 
  else
	return(true)
}

function stringValido(nombre){
  var checkOK = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZÄÁËÉÏÍÖÓÜÚ" + "abcdefghijklmnñopqrstuvwxyzäáëéïíöóüú +!$;.#-/_()*´`\\" + "0123456789";
  var valido = true; 
  for (i = 0; i < nombre.length; i++) {
    ch = nombre.charAt(i); 
    for (j = 0; j < checkOK.length; j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length) { 
      valido = false; 
      break; 
    }
  }
  if (!valido) 
	return (false); 
  else
	return(true)
}

function rutaArchivoValido(nombre){
  var checkOK = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZÄÁËÉÏÍÖÓÜÚ" + "abcdefghijklmnñopqrstuvwxyzäáëéïíöóüú +!$;.#-_()´`:\\" + "0123456789";
  var valido = true; 
  for (i = 0; i < nombre.length; i++) {
    ch = nombre.charAt(i); 
    for (j = 0; j < checkOK.length; j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length) { 
      valido = false; 
      break; 
    }
  }
  if (!valido) 
	return (false); 
  else
	return(true)
}

function ValidaCuit(nroCuit){
	//Descripcion: Valida que el cuil sea correcto
	// By lisandro Moro

	var valido = false;
	var Totalsuma;
	var Digito;
	var Resto;
	var Numerototal;
	var Numero1;
	var Numero2;
	var Numero3;
	var N1;
	var N2;
	var N3;
	var N4;
	var N5;
	var N6;
	var N7;
	var N8;
	var N9;
	var N10;
	var Opcion = "";

	Numerototal = nroCuit.toString();

	Numero1 = Numerototal.substr(0,2);
	Numero2 = Numerototal.substr(3,8);
	Numero3 = Numerototal.substr(12,1);
	
	N1 = Numero1.substr(0,1);
	N2 = Numero1.substr(1,1);
	
	N3 = Numero2.substr(0,1);
	N4 = Numero2.substr(1,1);
	N5 = Numero2.substr(2,1);
	N6 = Numero2.substr(3,1);
	N7 = Numero2.substr(4,1);
	N8 = Numero2.substr(5,1);
	N9 = Numero2.substr(6,1);
	N10 = Numero2.substr(7,1);

    if (Numerototal.length != 13){
        Opcion = "El número de CUIT está mal ingresado, debe contener trece caracteres.";
    }else{
        if (Numerototal.substr(2, 1) != "-"){
            Opcion = "El tercer carácter debe ser un guión.";
        }
        if (Numerototal.substr(11, 1) != "-" ){
            Opcion = "El decimosegundo carácter debe ser un guión.";
        }
        if (isNaN(Numero1)){
            Opcion = "Los dos primeros dígitos deben ser numéricos.";
        }
        if (isNaN(Numero2)){
            Opcion = "Los dígitos entre guiones deben ser numéricos.";
        }
        if (isNaN(Numero3)){
            Opcion = "El último dígito debe ser numérico.";
        }
  			
        Totalsuma = (N1*5) + (N2 * 4) + (N3 * 3) + (N4 * 2) + (N5 * 7) + (N6 * 6) + (N7 * 5) + (N8 * 4) + (N9 * 3) + (N10 * 2);
        Resto = Totalsuma % 11;

        switch(parseInt(Resto)){
	        case 0:
	            Digito = 0;
				break;
	        case 1:
	            Digito = 1;
				break;
	        default :
	            Digito = 11 - Resto;
				break;
        }
        if (Numero3 != Digito){
            Opcion = Opcion + " El Digito verificador es incorrecto.";
        }
	}
	if (Opcion == ""){
	    Valido = true;
	}else{
	    Valido = false;
		alert(Opcion);
	}
	return Valido;
}
function Left(str, n){
	if (n <= 0)
	    return "";
	else if (n > String(str).length)
	    return str;
	else
	    return String(str).substring(0,n);
}
function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}

