using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa12Acuerdos : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Acuerdos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ProyectoId = table.Column<int>(type: "int", nullable: false),
                    Numero = table.Column<int>(type: "int", nullable: false),
                    Titulo = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    RecursoQAId = table.Column<int>(type: "int", nullable: true),
                    ResponsableNombre = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: true),
                    FechaAcuerdo = table.Column<DateTime>(type: "datetime2", nullable: false),
                    FechaLimite = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Estado = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Acuerdos", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Acuerdos_Proyectos_ProyectoId",
                        column: x => x.ProyectoId,
                        principalTable: "Proyectos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                    table.ForeignKey(
                        name: "FK_Acuerdos_RecursosQA_RecursoQAId",
                        column: x => x.RecursoQAId,
                        principalTable: "RecursosQA",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateTable(
                name: "AcuerdoActividades",
                columns: table => new
                {
                    AcuerdoId = table.Column<int>(type: "int", nullable: false),
                    ActividadId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_AcuerdoActividades", x => new { x.AcuerdoId, x.ActividadId });
                    table.ForeignKey(
                        name: "FK_AcuerdoActividades_Actividades_ActividadId",
                        column: x => x.ActividadId,
                        principalTable: "Actividades",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                    table.ForeignKey(
                        name: "FK_AcuerdoActividades_Acuerdos_AcuerdoId",
                        column: x => x.AcuerdoId,
                        principalTable: "Acuerdos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_AcuerdoActividades_ActividadId",
                table: "AcuerdoActividades",
                column: "ActividadId");

            migrationBuilder.CreateIndex(
                name: "IX_Acuerdos_ProyectoId_Numero",
                table: "Acuerdos",
                columns: new[] { "ProyectoId", "Numero" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_Acuerdos_RecursoQAId",
                table: "Acuerdos",
                column: "RecursoQAId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "AcuerdoActividades");

            migrationBuilder.DropTable(
                name: "Acuerdos");
        }
    }
}
