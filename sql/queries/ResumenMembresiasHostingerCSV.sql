-- ==========================================================
-- Consulta : ResumenMembresiasHostingerCSV
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT codigolocal() AS sucursal, [ClientesClub]![CodCliente] AS membresia, [ClientesClub]![Nombre] AS nombre, [ClientesClub]![Apellidos] AS apellido, [ClientesClub]![Celular] AS celular, cfechasqlfecha([ClientesClub]![Cumpleanos]) AS fecha_nacimiento, [ClientesClub]![Correo] AS correo, cfechasqlfecha([ClientesClub]![Fecha de Inscripcion]) AS fecha_registro, [ClientesClub]![PuntosIniciales] AS puntos_iniciales, nombrelocalglobal(codigolocal()) AS nombre_sucursal
FROM ClientesClub;

