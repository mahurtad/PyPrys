using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa10TipoAusencia : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "Tipo",
                table: "Ausencias",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Tipo",
                table: "Ausencias");
        }
    }
}
