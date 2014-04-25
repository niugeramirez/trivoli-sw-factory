-- Update Recursos Descripcion con Nombre - Apellido Usuario
update recursosreservables
set descripcion = (select concat(nombre, ' ', apellido) from   usuarios where usuarios.id = recursosreservables.idusuario)
where recursosreservables.id > 0;