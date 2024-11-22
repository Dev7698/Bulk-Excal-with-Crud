using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace CrudUsingAjax.Migrations
{
    /// <inheritdoc />
    public partial class AddViewModel : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id1",
                table: "EmployeesDB");

            migrationBuilder.DropPrimaryKey(
                name: "PK_EmployeesDB",
                table: "EmployeesDB");

            migrationBuilder.DropPrimaryKey(
                name: "PK_DepartmentDB",
                table: "DepartmentDB");

            migrationBuilder.RenameTable(
                name: "EmployeesDB",
                newName: "Employee");

            migrationBuilder.RenameTable(
                name: "DepartmentDB",
                newName: "Department");

            migrationBuilder.RenameIndex(
                name: "IX_EmployeesDB_Department_Id1",
                table: "Employee",
                newName: "IX_Employee_Department_Id1");

            migrationBuilder.AddPrimaryKey(
                name: "PK_Employee",
                table: "Employee",
                column: "Employee_Id");

            migrationBuilder.AddPrimaryKey(
                name: "PK_Department",
                table: "Department",
                column: "Department_Id");

            migrationBuilder.CreateTable(
                name: "ViewEmployeeModel",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Employee_Id = table.Column<int>(type: "int", nullable: true),
                    Department_Id = table.Column<int>(type: "int", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ViewEmployeeModel", x => x.Id);
                    table.ForeignKey(
                        name: "FK_ViewEmployeeModel_Department_Department_Id",
                        column: x => x.Department_Id,
                        principalTable: "Department",
                        principalColumn: "Department_Id");
                    table.ForeignKey(
                        name: "FK_ViewEmployeeModel_Employee_Employee_Id",
                        column: x => x.Employee_Id,
                        principalTable: "Employee",
                        principalColumn: "Employee_Id");
                });

            migrationBuilder.CreateIndex(
                name: "IX_ViewEmployeeModel_Department_Id",
                table: "ViewEmployeeModel",
                column: "Department_Id");

            migrationBuilder.CreateIndex(
                name: "IX_ViewEmployeeModel_Employee_Id",
                table: "ViewEmployeeModel",
                column: "Employee_Id");

            migrationBuilder.AddForeignKey(
                name: "FK_Employee_Department_Department_Id1",
                table: "Employee",
                column: "Department_Id1",
                principalTable: "Department",
                principalColumn: "Department_Id");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Employee_Department_Department_Id1",
                table: "Employee");

            migrationBuilder.DropTable(
                name: "ViewEmployeeModel");

            migrationBuilder.DropPrimaryKey(
                name: "PK_Employee",
                table: "Employee");

            migrationBuilder.DropPrimaryKey(
                name: "PK_Department",
                table: "Department");

            migrationBuilder.RenameTable(
                name: "Employee",
                newName: "EmployeesDB");

            migrationBuilder.RenameTable(
                name: "Department",
                newName: "DepartmentDB");

            migrationBuilder.RenameIndex(
                name: "IX_Employee_Department_Id1",
                table: "EmployeesDB",
                newName: "IX_EmployeesDB_Department_Id1");

            migrationBuilder.AddPrimaryKey(
                name: "PK_EmployeesDB",
                table: "EmployeesDB",
                column: "Employee_Id");

            migrationBuilder.AddPrimaryKey(
                name: "PK_DepartmentDB",
                table: "DepartmentDB",
                column: "Department_Id");

            migrationBuilder.AddForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id1",
                table: "EmployeesDB",
                column: "Department_Id1",
                principalTable: "DepartmentDB",
                principalColumn: "Department_Id");
        }
    }
}
