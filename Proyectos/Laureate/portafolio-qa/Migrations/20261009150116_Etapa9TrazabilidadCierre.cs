using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa9TrazabilidadCierre : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<DateTime>(
                name: "FechaFinOriginal",
                table: "Asignaciones",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<bool>(
                name: "NoEjecutada",
                table: "Asignaciones",
                type: "bit",
                nullable: false,
                defaultValue: false);

            migrationBuilder.CreateTable(
                name: "ActividadAuditorias",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ProyectoId = table.Column<int>(type: "int", nullable: false),
                    ActividadId = table.Column<int>(type: "int", nullable: true),
                    Actividad = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    Fecha = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Usuario = table.Column<string>(type: "nvarchar(256)", maxLength: 256, nullable: false),
                    Accion = table.Column<string>(type: "nvarchar(40)", maxLength: 40, nullable: false),
                    Detalle = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    ProyectoCerrado = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ActividadAuditorias", x => x.Id);
                    table.ForeignKey(
                        name: "FK_ActividadAuditorias_Proyectos_ProyectoId",
                        column: x => x.ProyectoId,
                        principalTable: "Proyectos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "Cierres",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ProyectoId = table.Column<int>(type: "int", nullable: false),
                    Fecha = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Usuario = table.Column<string>(type: "nvarchar(256)", maxLength: 256, nullable: false),
                    AvanceFinal = table.Column<decimal>(type: "decimal(5,2)", nullable: false),
                    TotalActividades = table.Column<int>(type: "int", nullable: false),
                    ActividadesCompletadas = table.Column<int>(type: "int", nullable: false),
                    Reconstruido = table.Column<bool>(type: "bit", nullable: false),
                    Vigente = table.Column<bool>(type: "bit", nullable: false),
                    FechaReapertura = table.Column<DateTime>(type: "datetime2", nullable: true),
                    UsuarioReapertura = table.Column<string>(type: "nvarchar(256)", maxLength: 256, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Cierres", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Cierres_Proyectos_ProyectoId",
                        column: x => x.ProyectoId,
                        principalTable: "Proyectos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "CierreActividades",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CierreProyectoRegistroId = table.Column<int>(type: "int", nullable: false),
                    Nombre = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    Responsable = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    Estado = table.Column<string>(type: "nvarchar(30)", maxLength: 30, nullable: false),
                    Avance = table.Column<decimal>(type: "decimal(5,2)", nullable: false),
                    FechaInicio = table.Column<DateTime>(type: "datetime2", nullable: false),
                    FechaFin = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_CierreActividades", x => x.Id);
                    table.ForeignKey(
                        name: "FK_CierreActividades_Cierres_CierreProyectoRegistroId",
                        column: x => x.CierreProyectoRegistroId,
                        principalTable: "Cierres",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "CierreParticipantes",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CierreProyectoRegistroId = table.Column<int>(type: "int", nullable: false),
                    RecursoQAId = table.Column<int>(type: "int", nullable: false),
                    NombreRecurso = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    Rol = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: false),
                    Porcentaje = table.Column<decimal>(type: "decimal(5,2)", nullable: false),
                    FechaInicio = table.Column<DateTime>(type: "datetime2", nullable: false),
                    FechaFin = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Ejecutada = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_CierreParticipantes", x => x.Id);
                    table.ForeignKey(
                        name: "FK_CierreParticipantes_Cierres_CierreProyectoRegistroId",
                        column: x => x.CierreProyectoRegistroId,
                        principalTable: "Cierres",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                    table.ForeignKey(
                        name: "FK_CierreParticipantes_RecursosQA_RecursoQAId",
                        column: x => x.RecursoQAId,
                        principalTable: "RecursosQA",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateIndex(
                name: "IX_ActividadAuditorias_ProyectoId_Fecha",
                table: "ActividadAuditorias",
                columns: new[] { "ProyectoId", "Fecha" });

            migrationBuilder.CreateIndex(
                name: "IX_CierreActividades_CierreProyectoRegistroId",
                table: "CierreActividades",
                column: "CierreProyectoRegistroId");

            migrationBuilder.CreateIndex(
                name: "IX_CierreParticipantes_CierreProyectoRegistroId",
                table: "CierreParticipantes",
                column: "CierreProyectoRegistroId");

            migrationBuilder.CreateIndex(
                name: "IX_CierreParticipantes_RecursoQAId",
                table: "CierreParticipantes",
                column: "RecursoQAId");

            migrationBuilder.CreateIndex(
                name: "IX_Cierres_ProyectoId",
                table: "Cierres",
                column: "ProyectoId");

            // Proyectos ya cerrados: se reconstruye el historial posible con las asignaciones que sobrevivieron al cierre.
            // Las actividades de esos proyectos ya perdieron su responsable y no se puede recuperar.
            migrationBuilder.Sql(@"
INSERT INTO Cierres (ProyectoId, Fecha, Usuario, AvanceFinal, TotalActividades, ActividadesCompletadas, Reconstruido, Vigente)
SELECT p.Id,
       ISNULL(DATEADD(DAY, 1, (SELECT MAX(a.FechaFin) FROM Asignaciones a WHERE a.ProyectoId = p.Id)), p.FechaFin),
       N'Reconstruido (migración)', 100,
       (SELECT COUNT(*) FROM Actividades x WHERE x.ProyectoId = p.Id),
       (SELECT COUNT(*) FROM Actividades x WHERE x.ProyectoId = p.Id AND x.Estado = 2),
       1, 1
FROM Proyectos p
WHERE p.Estado = 3 AND NOT EXISTS (SELECT 1 FROM Cierres c WHERE c.ProyectoId = p.Id);");

            migrationBuilder.Sql(@"
INSERT INTO CierreParticipantes (CierreProyectoRegistroId, RecursoQAId, NombreRecurso, Rol, Porcentaje, FechaInicio, FechaFin, Ejecutada)
SELECT c.Id, a.RecursoQAId, LTRIM(RTRIM(r.Nombres + N' ' + r.Apellidos)), a.Rol, a.Porcentaje, a.FechaInicio, a.FechaFin, 1
FROM Cierres c
JOIN Asignaciones a ON a.ProyectoId = c.ProyectoId
JOIN RecursosQA r ON r.Id = a.RecursoQAId
WHERE c.Reconstruido = 1;");

            migrationBuilder.Sql(@"
INSERT INTO CierreActividades (CierreProyectoRegistroId, Nombre, Responsable, Estado, Avance, FechaInicio, FechaFin)
SELECT c.Id, x.Nombre, NULL,
       CASE x.Estado WHEN 0 THEN N'Pendiente' WHEN 1 THEN N'En curso' WHEN 2 THEN N'Completada' ELSE N'Cancelada' END,
       x.Avance, x.FechaInicio, x.FechaFin
FROM Cierres c
JOIN Actividades x ON x.ProyectoId = c.ProyectoId
WHERE c.Reconstruido = 1;");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ActividadAuditorias");

            migrationBuilder.DropTable(
                name: "CierreActividades");

            migrationBuilder.DropTable(
                name: "CierreParticipantes");

            migrationBuilder.DropTable(
                name: "Cierres");

            migrationBuilder.DropColumn(
                name: "FechaFinOriginal",
                table: "Asignaciones");

            migrationBuilder.DropColumn(
                name: "NoEjecutada",
                table: "Asignaciones");
        }
    }
}
