var SeparadorDecimal = ".";

// Descripción: Valida que el numero tenga X digitos enteros e Y digitos decimales 				
function validanumero(numero, x, y){
var numsplit, numstring;
var z;

if (isNaN(numero.value))
	return false;
else{
	k = "1" + "e+" + x
	limite = parseFloat(k)
	if (parseFloat(numero.value) > limite || parseFloat(numero.value) < -limite)
		return(false)
	else{ 
		numstring = numero.value.toString();
		numsplit = numstring.split(SeparadorDecimal);
		if (numero.value < 0)
			z = x + 1;
		else
			z = x;
		if (numsplit[0].length > z)
			return false; 
		else{
			if (numsplit.length == 1)
				return true;
			else{
				if ((numsplit[1].length <= y) && (numsplit[1].length > 0))
					return true;
				else
					return false;
				}
			}
		}	
	}
}

// Descripción: Redondea Num con Places decimales
function roundit(Num, Places) {
   if (Places > 0) {
      if ((Num.toString().length - Num.toString().lastIndexOf(SeparadorDecimal)) > (Places + 1)) {
         var Rounder = Math.pow(10, Places);
         return Math.round(Num * Rounder) / Rounder;
      }
      else return Num;
   }
   else return Math.round(Num);
}

