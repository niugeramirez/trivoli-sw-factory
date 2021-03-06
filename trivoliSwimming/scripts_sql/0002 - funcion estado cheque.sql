CREATE FUNCTION [dbo].[get_estado_cheque]
(
	-- Add the parameters for the function here
	@idcheque int
)
RETURNS int
AS
BEGIN
	-- Declare the return variable here
	DECLARE @ResultVar int;
	
	-- Add the T-SQL statements to compute the return value here
	SELECT @ResultVar= 	
		case   
			when cheques.flag_propio = -1 then  
				case 
					/*Chequeo que este asociado a una compra*/ 
					when (	select COUNT(*) 
							from	cajaMovimientos 
							inner join tiposMovimientoCaja on  tiposMovimientoCaja.id = cajaMovimientos.idtipoMovimiento 
							where	tiposMovimientoCaja.flagCompra = -1 
							and		cajaMovimientos.idcheque = cheques .id		
							) > 0 then  
									case 
										when ISNULL(cheques.flag_cobrado_pagado , 0) = -1 then 1--'PAGADO' 
										else 2--'ENTREGADO' 
									end 			
					else 3--'PENDIENTE ENTREGAR'  
					end 
			else  
				case  
					when (	/*chequeo si esta asociado a una venta*/ 
							select COUNT(*) 
							from	cajaMovimientos 
							inner join tiposMovimientoCaja on  tiposMovimientoCaja.id = cajaMovimientos.idtipoMovimiento 
							where	tiposMovimientoCaja.flagVenta = -1 
							and		cajaMovimientos.idcheque = cheques.id	 
							) > 0 then 
									case 
										/*Chequeo que este asociado a una compra*/ 
										when (	select COUNT(*) 
												from	cajaMovimientos 
												inner join tiposMovimientoCaja on  tiposMovimientoCaja.id = cajaMovimientos.idtipoMovimiento 
												where	tiposMovimientoCaja.flagCompra = -1 
												and		cajaMovimientos.idcheque = cheques.id						 
												) > 0 then 2--'ENTREGADO' 
										else  
															case 
										            			when ISNULL(cheques.flag_cobrado_pagado , 0) = -1 then 4--'COBRADO' 
																else 5--'PENDIENTE ENTREGAR/COBRAR' 
															end	 
									end		 	
					else 6--'PENDIENTE ASOCIAR VENTA' 
				end 
		end     
	FROM cheques 
	where cheques.id = @idcheque;
	
	-- Return the result of the function
	RETURN @ResultVar;

END

