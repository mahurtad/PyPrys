using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa14SquashProyectoVinculado : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "FuenteAvance",
                table: "Actividades");

            migrationBuilder.AddColumn<string>(
                name: "SquashLogin",
                table: "RecursosQA",
                type: "nvarchar(100)",
                maxLength: 100,
                nullable: true);

            migrationBuilder.AddColumn<int>(
                name: "SquashProjectId",
                table: "Proyectos",
                type: "int",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "SquashSincronizadoEn",
                table: "Proyectos",
                type: "datetime2",
                nullable: true);

            migrationBuilder.CreateTable(
                name: "SquashCampanas",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ProyectoId = table.Column<int>(type: "int", nullable: false),
                    SquashId = table.Column<int>(type: "int", nullable: false),
                    Nombre = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    Estado = table.Column<string>(type: "nvarchar(30)", maxLength: 30, nullable: true),
                    Vigente = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SquashCampanas", x => x.Id);
                    table.ForeignKey(
                        name: "FK_SquashCampanas_Proyectos_ProyectoId",
                        column: x => x.ProyectoId,
                        principalTable: "Proyectos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "SquashIteraciones",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    SquashCampanaId = table.Column<int>(type: "int", nullable: false),
                    SquashId = table.Column<int>(type: "int", nullable: false),
                    Nombre = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    Estado = table.Column<string>(type: "nvarchar(30)", maxLength: 30, nullable: true),
                    Porcentaje = table.Column<decimal>(type: "decimal(5,2)", precision: 5, scale: 2, nullable: false),
                    Vigente = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SquashIteraciones", x => x.Id);
                    table.ForeignKey(
                        name: "FK_SquashIteraciones_SquashCampanas_SquashCampanaId",
                        column: x => x.SquashCampanaId,
                        principalTable: "SquashCampanas",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "SquashCasos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    SquashIteracionId = table.Column<int>(type: "int", nullable: false),
                    SquashItemId = table.Column<int>(type: "int", nullable: false),
                    Nombre = table.Column<string>(type: "nvarchar(300)", maxLength: 300, nullable: false),
                    EstadoSquash = table.Column<string>(type: "nvarchar(30)", maxLength: 30, nullable: false),
                    Porcentaje = table.Column<decimal>(type: "decimal(5,2)", precision: 5, scale: 2, nullable: true),
                    AsignadoLogin = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    RecursoQAId = table.Column<int>(type: "int", nullable: true),
                    UltimaEjecucion = table.Column<DateTime>(type: "datetime2", nullable: true),
                    Vigente = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SquashCasos", x => x.Id);
                    table.ForeignKey(
                        name: "FK_SquashCasos_RecursosQA_RecursoQAId",
                        column: x => x.RecursoQAId,
                        principalTable: "RecursosQA",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                    table.ForeignKey(
                        name: "FK_SquashCasos_SquashIteraciones_SquashIteracionId",
                        column: x => x.SquashIteracionId,
                        principalTable: "SquashIteraciones",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            // Los recursos existentes usan su correo como login de Squash TM hasta que se ajuste en Administración → Recursos QA.
            migrationBuilder.Sql("UPDATE RecursosQA SET SquashLogin = Correo WHERE SquashLogin IS NULL");

            migrationBuilder.CreateIndex(
                name: "IX_RecursosQA_SquashLogin",
                table: "RecursosQA",
                column: "SquashLogin",
                unique: true,
                filter: "[SquashLogin] IS NOT NULL");

            migrationBuilder.CreateIndex(
                name: "IX_Proyectos_SquashProjectId",
                table: "Proyectos",
                column: "SquashProjectId",
                unique: true,
                filter: "[SquashProjectId] IS NOT NULL");

            migrationBuilder.CreateIndex(
                name: "IX_SquashCampanas_ProyectoId_SquashId",
                table: "SquashCampanas",
                columns: new[] { "ProyectoId", "SquashId" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_SquashCasos_RecursoQAId",
                table: "SquashCasos",
                column: "RecursoQAId");

            migrationBuilder.CreateIndex(
                name: "IX_SquashCasos_SquashIteracionId_SquashItemId",
                table: "SquashCasos",
                columns: new[] { "SquashIteracionId", "SquashItemId" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_SquashIteraciones_SquashCampanaId_SquashId",
                table: "SquashIteraciones",
                columns: new[] { "SquashCampanaId", "SquashId" },
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "SquashCasos");

            migrationBuilder.DropTable(
                name: "SquashIteraciones");

            migrationBuilder.DropTable(
                name: "SquashCampanas");

            migrationBuilder.DropIndex(
                name: "IX_RecursosQA_SquashLogin",
                table: "RecursosQA");

            migrationBuilder.DropIndex(
                name: "IX_Proyectos_SquashProjectId",
                table: "Proyectos");

            migrationBuilder.DropColumn(
                name: "SquashLogin",
                table: "RecursosQA");

            migrationBuilder.DropColumn(
                name: "SquashProjectId",
                table: "Proyectos");

            migrationBuilder.DropColumn(
                name: "SquashSincronizadoEn",
                table: "Proyectos");

            migrationBuilder.AddColumn<int>(
                name: "FuenteAvance",
                table: "Actividades",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }
    }
}
