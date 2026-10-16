using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa17VinculoBugzilla : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "VinculosBugzilla",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ProyectoId = table.Column<int>(type: "int", nullable: false),
                    SquashAmbiente = table.Column<string>(type: "nvarchar(300)", maxLength: 300, nullable: false),
                    SquashProjectId = table.Column<int>(type: "int", nullable: false),
                    BugzillaAmbiente = table.Column<string>(type: "nvarchar(300)", maxLength: 300, nullable: false),
                    ProductoId = table.Column<int>(type: "int", nullable: true),
                    ProductoNombre = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    ComponenteId = table.Column<int>(type: "int", nullable: true),
                    ComponenteNombre = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    VersionId = table.Column<int>(type: "int", nullable: true),
                    VersionNombre = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Estado = table.Column<int>(type: "int", nullable: false),
                    UltimaSincronizacion = table.Column<DateTime>(type: "datetime2", nullable: true),
                    UltimoIntento = table.Column<DateTime>(type: "datetime2", nullable: true),
                    UltimoError = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    AsociacionExistenteConfirmada = table.Column<bool>(type: "bit", nullable: false),
                    BloqueadoHasta = table.Column<DateTime>(type: "datetime2", nullable: true),
                    CreadoEn = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_VinculosBugzilla", x => x.Id);
                    table.ForeignKey(
                        name: "FK_VinculosBugzilla_Proyectos_ProyectoId",
                        column: x => x.ProyectoId,
                        principalTable: "Proyectos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_VinculosBugzilla_BugzillaAmbiente_ProductoId",
                table: "VinculosBugzilla",
                columns: new[] { "BugzillaAmbiente", "ProductoId" },
                unique: true,
                filter: "[ProductoId] IS NOT NULL");

            migrationBuilder.CreateIndex(
                name: "IX_VinculosBugzilla_ProyectoId",
                table: "VinculosBugzilla",
                column: "ProyectoId",
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_VinculosBugzilla_SquashAmbiente_SquashProjectId",
                table: "VinculosBugzilla",
                columns: new[] { "SquashAmbiente", "SquashProjectId" },
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "VinculosBugzilla");
        }
    }
}
