using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PortafolioQA.Migrations
{
    /// <inheritdoc />
    public partial class Etapa11ActividadesPredeterminadas : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<decimal>(
                name: "AvanceLegado",
                table: "Proyectos",
                type: "decimal(5,2)",
                nullable: true);

            migrationBuilder.AddColumn<bool>(
                name: "AvancePendienteConfirmar",
                table: "Proyectos",
                type: "bit",
                nullable: false,
                defaultValue: false);

            migrationBuilder.AddColumn<bool>(
                name: "EsPredeterminada",
                table: "Actividades",
                type: "bit",
                nullable: false,
                defaultValue: false);

            migrationBuilder.AddColumn<int>(
                name: "Orden",
                table: "Actividades",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<decimal>(
                name: "Peso",
                table: "Actividades",
                type: "decimal(5,2)",
                nullable: false,
                defaultValue: 10m);
            // Solo los proyectos sin actividades requieren tratamiento del avance heredado.
            migrationBuilder.Sql("""
                SELECT p.Id, p.Estado, p.Avance, CASE WHEN EXISTS
                    (SELECT 1 FROM Actividades a WHERE a.ProyectoId=p.Id) THEN 1 ELSE 0 END ConActividades
                INTO #ProyectosEtapa11 FROM Proyectos p;

                UPDATE p SET AvancePendienteConfirmar=1, AvanceLegado=p.Avance
                FROM Proyectos p JOIN #ProyectosEtapa11 b ON b.Id=p.Id
                WHERE b.ConActividades=0 AND b.Estado=2;

                DECLARE @Def TABLE (Orden int, Nombre nvarchar(200), Peso decimal(5,2), Acumulado int);
                INSERT INTO @Def VALUES
                (1,N'Diseño y redacción de Casos de Prueba',30,30),
                (2,N'Ejecución de Casos de prueba',50,80),
                (3,N'Elaborar informe',20,100);

                -- Si hubiera nombres duplicados, se protege únicamente la primera coincidencia.
                ;WITH Coincidencias AS (
                    SELECT a.Id,d.Nombre,d.Orden,d.Peso,ROW_NUMBER() OVER(PARTITION BY a.ProyectoId,d.Orden ORDER BY a.Id) rn
                    FROM Actividades a JOIN @Def d ON LTRIM(RTRIM(a.Nombre))=d.Nombre
                )
                UPDATE a SET Nombre=c.Nombre,EsPredeterminada=1,Orden=c.Orden,Peso=c.Peso
                FROM Actividades a JOIN Coincidencias c ON c.Id=a.Id WHERE c.rn=1;
                UPDATE Actividades SET Orden=4 WHERE EsPredeterminada=0;
                UPDATE Actividades SET Estado=CASE WHEN Avance>=100 THEN 2 WHEN Avance>0 THEN 1 ELSE 0 END WHERE Estado<>3;

                ;WITH Rangos AS (
                    SELECT p.Id,p.FechaInicio,p.FechaFin,b.Estado,b.ConActividades,
                        CASE WHEN DATEDIFF(day,p.FechaInicio,p.FechaFin)<0 THEN 1
                             ELSE DATEDIFF(day,p.FechaInicio,p.FechaFin)+1 END Dias
                    FROM Proyectos p JOIN #ProyectosEtapa11 b ON b.Id=p.Id
                ), Cortes AS (
                    SELECT *,CASE WHEN Dias>=3 THEN
                        CASE WHEN ROUND(Dias*0.3,0)<1 THEN 1 WHEN ROUND(Dias*0.3,0)>Dias-2 THEN Dias-2
                             ELSE CONVERT(int,ROUND(Dias*0.3,0)) END
                        ELSE CONVERT(int,CEILING(Dias*0.3)) END C1 FROM Rangos
                ), Limites AS (
                    SELECT *,CASE WHEN Dias>=3 THEN
                        CASE WHEN ROUND(Dias*0.8,0)<C1+1 THEN C1+1 WHEN ROUND(Dias*0.8,0)>Dias-1 THEN Dias-1
                             ELSE CONVERT(int,ROUND(Dias*0.8,0)) END
                        ELSE CONVERT(int,CEILING(Dias*0.8)) END C2 FROM Cortes
                )
                INSERT INTO Actividades (ProyectoId,Nombre,FechaInicio,FechaFin,Avance,Estado,RecursoQAId,EsPredeterminada,Orden,Peso)
                SELECT p.Id,d.Nombre,
                    DATEADD(day,CASE d.Orden WHEN 1 THEN 0 WHEN 2 THEN CASE WHEN C1>=Dias THEN Dias-1 ELSE C1 END
                        ELSE CASE WHEN C2>=Dias THEN Dias-1 ELSE C2 END END,CONVERT(date,p.FechaInicio)),
                    DATEADD(day,CASE d.Orden WHEN 1 THEN C1-1 WHEN 2 THEN C2-1 ELSE Dias-1 END,CONVERT(date,p.FechaInicio)),
                    CASE WHEN ConActividades=0 AND p.Estado=3 THEN 100 ELSE 0 END,
                    CASE WHEN ConActividades=0 AND p.Estado=3 THEN 2 WHEN ConActividades=0 AND p.Estado=4 THEN 3 ELSE 0 END,
                    NULL,1,d.Orden,d.Peso
                FROM Limites p CROSS JOIN @Def d
                WHERE NOT EXISTS (SELECT 1 FROM Actividades a WHERE a.ProyectoId=p.Id AND a.EsPredeterminada=1 AND a.Orden=d.Orden);

                -- El estado y sus efectos (foto de cierre / reapertura) se reconcilian con CierreProyecto
                -- al iniciar la aplicación o con el ejecutor de verificación, antes de reparar coherencia.
                UPDATE p SET Avance=CASE WHEN x.Promedio>=100 THEN 100
                    WHEN ROUND(x.Promedio,2)>=100 THEN 99.99 ELSE ROUND(x.Promedio,2) END
                FROM Proyectos p JOIN #ProyectosEtapa11 b ON p.Id=b.Id
                CROSS APPLY (SELECT COALESCE(SUM(a.Avance*a.Peso)/NULLIF(SUM(a.Peso),0),0) Promedio
                    FROM Actividades a WHERE a.ProyectoId=p.Id AND a.Estado<>3) x
                WHERE b.ConActividades=1 AND p.Estado<>4;
                DROP TABLE #ProyectosEtapa11;
                """);

        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "AvanceLegado",
                table: "Proyectos");

            migrationBuilder.DropColumn(
                name: "AvancePendienteConfirmar",
                table: "Proyectos");

            migrationBuilder.DropColumn(
                name: "EsPredeterminada",
                table: "Actividades");

            migrationBuilder.DropColumn(
                name: "Orden",
                table: "Actividades");

            migrationBuilder.DropColumn(
                name: "Peso",
                table: "Actividades");
        }
    }
}
