using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa18SquashModulos : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "SquashModulos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    SquashIteracionId = table.Column<int>(type: "int", nullable: false),
                    SquashId = table.Column<int>(type: "int", nullable: false),
                    Nombre = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    Vigente = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SquashModulos", x => x.Id);
                    table.ForeignKey(
                        name: "FK_SquashModulos_SquashIteraciones_SquashIteracionId",
                        column: x => x.SquashIteracionId,
                        principalTable: "SquashIteraciones",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "SquashModuloCaso",
                columns: table => new
                {
                    CasoId = table.Column<int>(type: "int", nullable: false),
                    ModuloId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SquashModuloCaso", x => new { x.CasoId, x.ModuloId });
                    table.ForeignKey(
                        name: "FK_SquashModuloCaso_SquashCasos_CasoId",
                        column: x => x.CasoId,
                        principalTable: "SquashCasos",
                        principalColumn: "Id");
                    table.ForeignKey(
                        name: "FK_SquashModuloCaso_SquashModulos_ModuloId",
                        column: x => x.ModuloId,
                        principalTable: "SquashModulos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_SquashModuloCaso_ModuloId",
                table: "SquashModuloCaso",
                column: "ModuloId");

            migrationBuilder.CreateIndex(
                name: "IX_SquashModulos_SquashIteracionId_SquashId",
                table: "SquashModulos",
                columns: new[] { "SquashIteracionId", "SquashId" },
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "SquashModuloCaso");

            migrationBuilder.DropTable(
                name: "SquashModulos");
        }
    }
}
