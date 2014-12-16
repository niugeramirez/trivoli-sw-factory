function h_correcta(hora,minutos)
{
if (isNaN(hora)|| (hora<0)|| (hora>23) || isNaN(minutos)|| (minutos<0)|| (minutos>59)
		|| (hora.length != 2) || (minutos.length != 2)) 
		if ((hora=='24')&&(minutos=='00'))
		  	return true;
		else
			return false;	
return true;
}

function h_esmenor(hora1,minutos1,hora2,minutos2){
	if (hora1 > hora2){
		return false;
	}
	else
		if (hora1 == hora2){
			if (minutos1 >= minutos2){
				return false;
				}
		}
	return true	;
}


