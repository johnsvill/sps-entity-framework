using Microsoft.EntityFrameworkCore.Migrations;

namespace Procs.Data.Migrations
{
    public partial class ProcedureProductos : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            var procedure = @"CREATE PROCEDURE [prueba].sp_Productos
            AS
            BEGIN
            SELECT * FROM [prueba].[Productos]
            END
            ";

            migrationBuilder.Sql(procedure);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            var procedure = @"DROP PROCEDURE [prueba].sp_Productos";

            migrationBuilder.Sql(procedure);
        }
    }
}
