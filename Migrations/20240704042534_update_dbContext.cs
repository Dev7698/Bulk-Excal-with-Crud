using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace CrudUsingAjax.Migrations
{
    /// <inheritdoc />
    public partial class update_dbContext : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id",
                table: "EmployeesDB");

            migrationBuilder.AddForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id",
                table: "EmployeesDB",
                column: "Department_Id",
                principalTable: "DepartmentDB",
                principalColumn: "Department_Id",
                onDelete: ReferentialAction.Restrict);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id",
                table: "EmployeesDB");

            migrationBuilder.AddForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id",
                table: "EmployeesDB",
                column: "Department_Id",
                principalTable: "DepartmentDB",
                principalColumn: "Department_Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}
