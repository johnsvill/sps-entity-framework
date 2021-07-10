using Microsoft.EntityFrameworkCore.Migrations;

namespace Procs.Data.Migrations
{
    public partial class PrimeraMigracion : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
                var procedure = @" CREATE PROCEDURE prueba._sp_NewId
                AS 
                BEGIN
                SELECT NEWID() AS ID
                END ";

            migrationBuilder.Sql(procedure);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql("DROP PROCEDURE prueba._sp_NewId");
        }
    }
}
