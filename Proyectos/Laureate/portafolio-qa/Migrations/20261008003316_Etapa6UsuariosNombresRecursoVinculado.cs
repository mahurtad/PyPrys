using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa6UsuariosNombresRecursoVinculado : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "UsuarioId",
                table: "RecursosQA",
                type: "nvarchar(450)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "Apellidos",
                table: "AspNetUsers",
                type: "nvarchar(100)",
                maxLength: 100,
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Nombres",
                table: "AspNetUsers",
                type: "nvarchar(100)",
                maxLength: 100,
                nullable: false,
                defaultValue: "");

            // Conserva los datos: el primer token del nombre completo pasa a Nombres y el resto a Apellidos.
            migrationBuilder.Sql(@"
UPDATE AspNetUsers SET
    Nombres   = CASE WHEN CHARINDEX(' ', LTRIM(RTRIM(NombreCompleto))) > 0
                     THEN LEFT(LTRIM(RTRIM(NombreCompleto)), CHARINDEX(' ', LTRIM(RTRIM(NombreCompleto))) - 1)
                     ELSE LTRIM(RTRIM(NombreCompleto)) END,
    Apellidos = CASE WHEN CHARINDEX(' ', LTRIM(RTRIM(NombreCompleto))) > 0
                     THEN LTRIM(SUBSTRING(LTRIM(RTRIM(NombreCompleto)), CHARINDEX(' ', LTRIM(RTRIM(NombreCompleto))) + 1, 200))
                     ELSE '' END;");

            migrationBuilder.DropColumn(
                name: "NombreCompleto",
                table: "AspNetUsers");

            migrationBuilder.CreateIndex(
                name: "IX_RecursosQA_UsuarioId",
                table: "RecursosQA",
                column: "UsuarioId",
                unique: true,
                filter: "[UsuarioId] IS NOT NULL");

            migrationBuilder.AddForeignKey(
                name: "FK_RecursosQA_AspNetUsers_UsuarioId",
                table: "RecursosQA",
                column: "UsuarioId",
                principalTable: "AspNetUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            // Vincula los recursos existentes con el usuario de igual correo (si lo hay).
            migrationBuilder.Sql(@"
UPDATE r SET r.UsuarioId = u.Id
FROM RecursosQA r
JOIN AspNetUsers u ON u.NormalizedEmail = UPPER(r.Correo)
WHERE r.UsuarioId IS NULL;");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_RecursosQA_AspNetUsers_UsuarioId",
                table: "RecursosQA");

            migrationBuilder.DropIndex(
                name: "IX_RecursosQA_UsuarioId",
                table: "RecursosQA");

            migrationBuilder.DropColumn(
                name: "UsuarioId",
                table: "RecursosQA");

            migrationBuilder.DropColumn(
                name: "Apellidos",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "Nombres",
                table: "AspNetUsers");

            migrationBuilder.AddColumn<string>(
                name: "NombreCompleto",
                table: "AspNetUsers",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.Sql("UPDATE AspNetUsers SET NombreCompleto = LTRIM(RTRIM(Nombres + ' ' + Apellidos));");
        }
    }
}
