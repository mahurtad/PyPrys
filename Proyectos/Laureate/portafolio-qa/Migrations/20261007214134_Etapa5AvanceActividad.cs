using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa5AvanceActividad : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<decimal>(
                name: "Avance",
                table: "Actividades",
                type: "decimal(5,2)",
                nullable: false,
                defaultValue: 0m);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Avance",
                table: "Actividades");
        }
    }
}
