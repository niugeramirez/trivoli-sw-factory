	//Alvaro Bayon - 07/10/2003 
	//Las fechas admitidas son:
	   //		Con separador: cualquier formato válido dentro del formato español
	   //		Sin separador: ddmmaa o bien ddmmaaaa  
	   // Los separadores son "-" y "/"
	//Scarpa D. - 19/11/2003 	   
	   //  Se agregaron las funciones validarAnyo y validarMes
	
    /**
    * definimos las varables globales que van a contener la fecha completa, cada una de sus partes
    * y los dias correspondientes al mes de febrero segun sea el año bisiesto o no
    */
		/*
	FFavre - 12-01-04
	Se agrego la funcion Menor (estricto)
	Scarpa D. - 23-11-04
	Se agrego la funcion FechaDiff
	
	*/

    var a, mes, dia, anyo, febrero;
    

	function anyoCompleto(anyo)
    {
        /**
        * si el año introducido es de dos cifras lo pasamos a:
			[0,30]: Siglo XXI
			(30,99]: Siglo XX
        */
		
		var x = Number(anyo);
        if ((x < 100) && (x > 30))
            x = x + 1900;
        else
			if ((x > -1) && (x < 31))
	            x = x + 2000;
		return x;	
	}	
    /**
    * funcion para comprobar si una año es bisiesto
    * argumento anyo > año extraido de la fecha introducida por el usuario
    */
    function anyoBisiesto(anyo)
    {
        /**
        * si el año introducido es de dos cifras lo pasamos al periodo de 1900. Ejemplo: 25 > 1925
        */
		anyo = parseInt(anyo)
        if (anyo < 100)
            var fin = anyo + 2000;
        else
            var fin = anyo ;

		alert(fin);	
        /*
        * primera condicion: si el resto de dividir el año entre 4 no es cero > el año no es bisiesto
        * es decir, obtenemos año modulo 4, teniendo que cumplirse anyo mod(4)=0 para bisiesto
        */
        if (fin % 4 != 0)
            return false;
        else
        {
            if (fin % 100 == 0)
            {
                /**
                * si el año es divisible por 4 y por 100 y divisible por 400 > es bisiesto
                */
                if (fin % 400 == 0)
                {
                    return true;
                }
                /**
                * si es divisible por 4 y por 100 pero no lo es por 400 > no es bisiesto
                */
                else
                {
                    return false;
                }
            }
            /**
            * si es divisible por 4 y no es divisible por 100 > el año es bisiesto
            */
            else
            {
                return true;
            }
        }
    }
    
    /**
    * funcion principal de validacion de la fecha
    * argumento fecha > cadena de texto de la fecha introducida por el usuario
    */
    function validarfecha(fecha)
    {
       /**
       * obtenemos la fecha introducida y la separamos en dia, mes y año
       */
	   var a = fecha.value;
	  /* if (a.length != 10)
       {
           alert("El formato de la fecha debe ser DD/MM/AAAA. Por favor, introduzca un valor correcto");
           fecha.focus();
           fecha.select();
           return false;
       }
	   if ((a.substr(2,1) != "/")  ||
	       (a.substr(5,1) != "/"))   
       {
           alert("El formato de la fecha debe ser DD/MM/AAAA. Por favor, introduzca un valor correcto");
           fecha.focus();
           fecha.select();
           return false;
       }*/
   
		var conSeparador = false;
		// Se admiten / y - como separadores
   		if (a.indexOf("/",1)>0)
		{
	       dia=a.split("/")[0];
    	   mes=a.split("/")[1];
	       anyo=a.split("/")[2];
		   conSeparador = true;
      	}
   		if (a.indexOf("-",1)>0)
		{
	       dia=a.split("-")[0];
    	   mes=a.split("-")[1];
	       anyo=a.split("-")[2];
		   conSeparador = true;
      	}

		// Si no existe ningún separador válido permito cargar los formatos
		// ddmmaa o ddmmaaaa
		if (!conSeparador)
		{
			if (a.length == 6)
			{
				dia=a.substr(0,2);
				mes=a.substr(2,2);
				anyo=a.substr(4,2);
			}
			else
			if (a.length == 8)
			{
				dia=a.substr(0,2);
				mes=a.substr(2,2);
				anyo=a.substr(4,4);
			}
			else
			{
				alert("La fecha ingresada no es válida. Por favor, introduzca una fecha correcta.");
				return false;
			}
		}
		
	   alert(anyo);	
       if(anyoBisiesto(anyo))
           febrero=29;
       else
           febrero=28;
       /**
       * si el mes introducido es negativo, 0 o mayor que 12 > alertamos y detenemos ejecucion
       */
	   if (isNaN(dia) || isNaN(mes) || isNaN(anyo))
	   {
           alert("Por favor, ingrese unicamente dígitos para el dia, mes y año.");
           fecha.focus();
           fecha.select();
           return false;
       }
	   
       if ((mes<1) || (mes>12))
       {
           alert("El mes introducido no es válido. Por favor, introduzca un mes correcto");
           fecha.focus();
           fecha.select();
           return false;
       }
       /**
       * si el mes introducido es febrero y el dia es mayor que el correspondiente 
       * al año introducido > alertamos y detenemos ejecucion
       */
	   alert(mes);
	   alert(dia);
	   alert(febrero);
	   alert(((mes==2) && ((dia<1) || (dia>febrero))));
	   
       if ((mes==2) && ((dia<1) || (dia>febrero)))
       {
           alert("El dia introducido no es válido. Por favor, introduzca un dia correcto");
           fecha.focus();
           fecha.select();
           return false;
       }
       /**
       * si el mes introducido es de 31 dias y el dia introducido es mayor de 31 > alertamos y detenemos ejecucion
       */
       if (((mes==1) || (mes==3) || (mes==5) || (mes==7) || (mes==8) || (mes==10) || (mes==12)) && ((dia<1) || (dia>31)))
       {
           alert("El dia introducido no es válido. Por favor, introduzca un dia correcto");
           fecha.focus();
           fecha.select();
           return false;
       }
       /**
       * si el mes introducido es de 30 dias y el dia introducido es mayor de 301 > alertamos y detenemos ejecucion
       */
       if (((mes==4) || (mes==6) || (mes==9) || (mes==11)) && ((dia<1) || (dia>30)))
       {
           alert("El dia introducido no es válido. Por favor, introduzca un dia correcto");
           fecha.focus();
           fecha.select();
           return false;
       }
		
		anyo = anyoCompleto(anyo);
		//alert (anyo);
       /**
       * si el mes año introducido es menor que 1900 o mayor que 2010 > alertamos y detenemos ejecucion
       * NOTA: estos valores son a eleccion vuestra, y no constituyen por si solos fecha erronea
       */
       if ((anyo<1900) || (anyo>2100))
       {
           alert("El año introducido no es válido. Por favor, introduzca un año entre 1900 y 2100");
           fecha.focus();
           fecha.select();
		   return false;
       } 
       /**
       * en caso de que todo sea correcto > enviamos los datos del formulario
       * para ello debeis descomentar la ultima sentencia
       */
       else
         {fecha.value= comp2(dia)+"/"+comp2(mes)+"/"+anyo;
		 return true;
		 }
    }    

function comp2(num){
	if(num.length==1)
		num= "0"+num;
	return(num);	
}	
function consultafecha(actual){
/*esta funcion permite girar la fecha para poder usarla en consultas contra el Informix*/
  var auxi;
  auxi  = "'"+actual.substr(6,4) + "/" + actual.substr(3,2) + "/" + actual.substr(0,2)+"'";
  return auxi;
}

function cambiafecha(actual,ctf,base){
/*esta funcion permite girar la fecha para poder usarla en consultas contra el Informix*/
  var auxi;
  if (base ==1 || base ==5) 
  	auxi  = actual.substr(6,4) + "/" + actual.substr(3,2) + "/" + actual.substr(0,2);
  else	
  if ((base ==2) || (base ==8) || (base ==12) || (base ==13) || (base ==14) || (base ==15))
	//Usuario con idioma inglés
  	auxi  = actual.substr(3,2) + "/" + actual.substr(0,2) + "/" + actual.substr(6,4);
	//Usuario con idioma español
  	//auxi  = actual.substr(0,2) + "/" + actual.substr(3,2) + "/" + actual.substr(6,4);
  else{	
     if (base == 10){
         auxi  = actual.substr(3,2) + "/" + actual.substr(0,2) + "/" + actual.substr(6,4);
     }else{
  	  auxi = actual.substr(0,2) + "/" + actual.substr(3,2) + "/" + actual.substr(6,4);
     }
   }

  if (ctf)
  	auxi= "'"+auxi+"'";
  return auxi;
}

function menorque(fecha1,fecha2){
	var f1= new Date(); 
	f1.setFullYear(fecha1.substr(6,4),fecha1.substr(3,2)-1,fecha1.substr(0,2));
	var segf1=Date.parse(f1); 

	var f2= new Date(); 
	f2.setFullYear(fecha2.substr(6,4),fecha2.substr(3,2)-1,fecha2.substr(0,2));
	var segf2=Date.parse(f2); 

	if ((segf1<segf2)||(fecha1==fecha2)){return true}
	else{return false}
}

function menor(fecha1,fecha2){
	var f1= new Date(); 
	f1.setFullYear(fecha1.substr(6,4),fecha1.substr(3,2)-1,fecha1.substr(0,2));
	var segf1=Date.parse(f1); 

	var f2= new Date(); 
	f2.setFullYear(fecha2.substr(6,4),fecha2.substr(3,2)-1,fecha2.substr(0,2));
	var segf2=Date.parse(f2); 

	if (segf1<segf2){return true}
	else{return false}
}

function contardias( start, end) {

    var iOut = 0;
	var vstart = start.substr(3,2)+'/'+ start.substr(0,2) +'/'+ start.substr(6,4);
	var vend   = end.substr(3,2)  +'/'+ end.substr(0,2)+'/'+end.substr(6,4);
    
    var bufferA = Date.parse( vstart ) ;
    var bufferB = Date.parse( vend ) ;
    	
    var number = bufferB-bufferA ;
    
    iOut = parseInt(number / 86400000) ;
    iOut += parseInt((number % 86400000)/43200001) ;
    return iOut ;
}

function validarAnyo(obj){
  var a;
  var errores = 1;
  
  if (obj.value == ""){
    alert('Debe ingresar un año.');	
  }else{
    if (obj.value.length != 4){
	   if (obj.value.length < 3){
	      if (isNaN(obj.value)){
 	         alert('El año no es valido.');		  
		  }else{
	         a = parseInt(obj.value);
	         obj.value = 2000 + a;
		     errores = 0;
		  }
	   }else{
	      alert('El año no es valido.');
	   }
	}else{
	   if (isNaN(obj.value)){
	      alert('El año no es valido.');
	   }else{
	      a = parseInt(obj.value);
	      if ( (a <= 1900) || (a >= 2100) ){
	         alert("El año introducido no es válido. Por favor, introduzca un año entre 1900 y 2100");
		  }else{
		     errores = 0;
		  }
	   }		 
	}  
  }
  
  return (errores == 0);
}

function validarMes(obj){
  var a;
  var errores = 1;
  
  if (obj.value == ""){
    alert('Debe ingresar un mes.');	
  }else{
    if (obj.value.length < 3){
	   if (isNaN(obj.value)){
 	      alert('El mes no es valido.');		  
	   }else{
	      a = parseInt(obj.value);
	      if ( (a < 1) || (a > 12) ){
	         alert("El mes introducido no es válido. Por favor, introduzca un mes entre 1 y 12");
		  }else{	 
		     errores = 0;
		  }
	   }
	}else{
      alert('El mes no es valido.');
	}  
  }
  
  return (errores == 0);
}

function FechaDiff( desde, hasta, intervalo, redondear ) {

    var iOut = 0;
    var dia,mes,anno;
	var d1 = new Date();
	var d2 = new Date();	
	
	dia  = desde.substr(0,2);
	mes  = desde.substr(3,2)-1;
	anno = desde.substr(6,4);		
	d1.setFullYear(anno,mes,dia);

	dia  = hasta.substr(0,2);
	mes  = hasta.substr(3,2)-1;
	anno = hasta.substr(6,4);		
	d2.setFullYear(anno,mes,dia);	
   
    var numero = Date.parse(d2) - Date.parse(d1);;
    switch (intervalo.charAt(0))
    {
        case 'd': case 'D': 
            iOut = parseInt(numero / 86400000) ;
            if(redondear) iOut += parseInt((numero % 86400000)/43200001) ;
            break ;
        case 'h': case 'H':
            iOut = parseInt(numero / 3600000 ) ;
            if(redondear) iOut += parseInt((numero % 3600000)/1800001) ;
            break ;
        case 'm': case 'M':
            iOut = parseInt(numero / 60000 ) ;
            if(redondear) iOut += parseInt((numero % 60000)/30001) ;
            break ;
        case 's': case 'S':
            iOut = parseInt(numero / 1000 ) ;
            if(redondear) iOut += parseInt((numero % 1000)/501) ;
            break ;
        default:
        alert('Intervalo incorrecto.') ;
        return null ;
    }

    return iOut;
}


