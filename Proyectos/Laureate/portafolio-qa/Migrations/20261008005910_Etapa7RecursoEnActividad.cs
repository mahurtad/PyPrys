using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa7RecursoEnActividad : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "RecursoQAId",
                table: "Actividades",
                type: "int",
                nullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_Actividades_RecursoQAId",
                table: "Actividades",
                column: "RecursoQAId");

            migrationBuilder.AddForeignKey(
                name: "FK_Actividades_RecursosQA_RecursoQAId",
                table: "Actividades",
                column: "RecursoQAId",
                principalTable: "RecursosQA",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Actividades_RecursosQA_RecursoQAId",
                table: "Actividades");

            migrationBuilder.DropIndex(
                name: "IX_Actividades_RecursoQAId",
                table: "Actividades");

            migrationBuilder.DropColumn(
                name: "RecursoQAId",
                table: "Actividades");
        }
    }
}
