using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa8EstadoPorAvance : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            // Avance = promedio del avance de las actividades no canceladas (Estado 3 = Cancelada; 100 si todas están canceladas).
            migrationBuilder.Sql(@"
UPDATE p SET p.Avance = ISNULL((SELECT ROUND(AVG(a.Avance), 2) FROM Actividades a WHERE a.ProyectoId = p.Id AND a.Estado <> 3), 100)
FROM Proyectos p
WHERE p.Estado <> 4 AND EXISTS (SELECT 1 FROM Actividades a WHERE a.ProyectoId = p.Id);");

            // Se retira EnEjecucion (1); el estado se deriva del avance. Cancelado (4) no se toca.
            migrationBuilder.Sql(@"
UPDATE Proyectos SET Estado = CASE WHEN Avance >= 100 THEN 3 WHEN Avance <= 0 THEN 0 ELSE 2 END
WHERE Estado IN (0, 1, 2, 3);");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            // Irreversible a nivel de datos: no se conserva qué proyectos estaban En ejecución.
        }
    }
}
