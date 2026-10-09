using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa2Planificacion : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Ausencias",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    RecursoQAId = table.Column<int>(type: "int", nullable: false),
                    FechaInicio = table.Column<DateTime>(type: "datetime2", nullable: false),
                    FechaFin = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Motivo = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Ausencias", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Ausencias_RecursosQA_RecursoQAId",
                        column: x => x.RecursoQAId,
                        principalTable: "RecursosQA",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateTable(
                name: "Feriados",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Fecha = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Nombre = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Feriados", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Hitos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ProyectoId = table.Column<int>(type: "int", nullable: false),
                    Fecha = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Nombre = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Hitos", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Hitos_Proyectos_ProyectoId",
                        column: x => x.ProyectoId,
                        principalTable: "Proyectos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Ausencias_RecursoQAId",
                table: "Ausencias",
                column: "RecursoQAId");

            migrationBuilder.CreateIndex(
                name: "IX_Feriados_Fecha",
                table: "Feriados",
                column: "Fecha",
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_Hitos_ProyectoId",
                table: "Hitos",
                column: "ProyectoId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Ausencias");

            migrationBuilder.DropTable(
                name: "Feriados");

            migrationBuilder.DropTable(
                name: "Hitos");
        }
    }
}
