using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace CrudUsingAjax.Migrations
{
    /// <inheritdoc />
    public partial class UpdateDbContext : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id1",
                table: "EmployeesDB");

            migrationBuilder.DropIndex(
                name: "IX_EmployeesDB_Department_Id1",
                table: "EmployeesDB");

            migrationBuilder.DropColumn(
                name: "Department_Id1",
                table: "EmployeesDB");

            migrationBuilder.CreateIndex(
                name: "IX_EmployeesDB_Department_Id",
                table: "EmployeesDB",
                column: "Department_Id");

            migrationBuilder.AddForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id",
                table: "EmployeesDB",
                column: "Department_Id",
                principalTable: "DepartmentDB",
                principalColumn: "Department_Id",
                onDelete: ReferentialAction.Cascade);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id",
                table: "EmployeesDB");

            migrationBuilder.DropIndex(
                name: "IX_EmployeesDB_Department_Id",
                table: "EmployeesDB");

            migrationBuilder.AddColumn<int>(
                name: "Department_Id1",
                table: "EmployeesDB",
                type: "int",
                nullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_EmployeesDB_Department_Id1",
                table: "EmployeesDB",
                column: "Department_Id1");

            migrationBuilder.AddForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id1",
                table: "EmployeesDB",
                column: "Department_Id1",
                principalTable: "DepartmentDB",
                principalColumn: "Department_Id");
        }
    }
}
