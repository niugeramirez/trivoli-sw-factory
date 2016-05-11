--select * from tiposMovimientoCaja
update tiposMovimientoCaja set flagCompra = null
where tiposMovimientoCaja.descripcion like 'Invers%'