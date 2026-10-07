-- ==========================================================
-- Consulta : ClientesClubUnionExternoInterno
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT 
    codigolocal() AS sucursal,
    [ClientesClub]![CodCliente] AS membresia,
    [ClientesClub]![Nombre] AS nombre,
    [ClientesClub]![Apellidos] AS apellido,
    [ClientesClub]![Celular] AS celular,
    [ClientesClub]![Cumpleanos] AS fecha_nacimiento,
    [ClientesClub]![Correo] AS correo,
    [ClientesClub]![Fecha de Inscripcion] AS fecha_registro,
    [ClientesClub]![PuntosIniciales] AS puntos_iniciales,
    nombrelocalglobal(codigolocal()) AS nombre_sucursal,
    'Interno' AS origen  

FROM ClientesClub

UNION ALL SELECT 
    sucursal,
    membresia,
    nombre,
    apellido,
    celular,
    fecha_nacimiento,
    correo,
    fecha_registro,
    puntos_iniciales,
    nombre_sucursal,
    'Externo' AS origen

FROM clientesclubexterno;

