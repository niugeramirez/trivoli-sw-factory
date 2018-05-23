SELECT
  pagos.fecha,
  clientespacientes.apellido,
  clientespacientes.nombre, 
  practicas.descripcion,
  pagos.importe,
  recursosreservables.descripcion medico,
  pagos.idmediodepago,
  mediosdepago.titulo + ' ' + (CASE
    WHEN mediosdepago.flag_obrasocial = -1 THEN obrassociales.descripcion
    ELSE ' '
  END) titulo
FROM pagos
LEFT JOIN practicasrealizadas
  ON practicasrealizadas.id = pagos.idpracticarealizada
LEFT JOIN visitas
  ON visitas.id = practicasrealizadas.idvisita
LEFT JOIN clientespacientes
  ON clientespacientes.id = visitas.idpaciente
LEFT JOIN recursosreservables
  ON recursosreservables.id = visitas.idrecursoreservable
LEFT JOIN mediosdepago
  ON mediosdepago.id = pagos.idmediodepago
LEFT JOIN obrassociales
  ON obrassociales.id = clientespacientes.idobrasocial
LEFT JOIN practicas
  ON practicas.id = practicasrealizadas.idpractica
WHERE pagos.fecha >= '05/02/2018'
AND pagos.fecha <= '05/02/2018'
union
select   visitas.fecha,
  clientespacientes.apellido,
  clientespacientes.nombre, 
  practicas.descripcion,
  practicasrealizadas.precio -
  isnull( ( select sum(pagos.importe)
	from pagos
	where pagos.idpracticarealizada = practicasrealizadas.id
	and pagos.fecha >= '05/02/2018'
	AND pagos.fecha <= '05/02/2018'
  ),0) importe, 
  recursosreservables.descripcion medico,  
  99999999999999999 as idmediodepago,
  'SALDO DEUDOR: '+obrassociales.descripcion   titulo
from visitas
	inner join  practicasrealizadas ON visitas.id = practicasrealizadas.idvisita
	inner join  practicas			ON practicasrealizadas.idpractica = practicas.id
	inner JOIN clientespacientes	ON clientespacientes.id = visitas.idpaciente
	inner JOIN recursosreservables	ON recursosreservables.id = visitas.idrecursoreservable	
	LEFT JOIN obrassociales			ON obrassociales.id = clientespacientes.idobrasocial
where visitas.fecha >= '05/02/2018'
AND visitas.fecha <= '05/02/2018'
AND visitas.empnro = 1
and ISNULL(visitas.flag_ausencia,0) =0
and (
  practicasrealizadas.precio -
  isnull( ( select sum(pagos.importe)
	from pagos
	where pagos.idpracticarealizada = practicasrealizadas.id
	and pagos.fecha >= '05/02/2018'
	AND pagos.fecha <= '05/02/2018'
  ),0) <> 0
	)
ORDER BY  idmediodepago, titulo, fecha