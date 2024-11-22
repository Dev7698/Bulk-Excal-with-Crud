using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace CrudUsingAjax.Migrations
{
    /// <inheritdoc />
    public partial class EmployeeManagement : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "DepartmentDB",
                columns: table => new
                {
                    Department_Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DepartmentName = table.Column<int>(type: "int", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_DepartmentDB", x => x.Department_Id);
                });

            migrationBuilder.CreateTable(
                name: "ImageDB",
                columns: table => new
                {
                    Image_Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Name = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ImageDB", x => x.Image_Id);
                });

            migrationBuilder.CreateTable(
                name: "EmployeesDB",
                columns: table => new
                {
                    Employee_Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    First_Name = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Last_Name = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Email = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Phone_Number = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Gender = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Department_Id = table.Column<int>(type: "int", nullable: true),
                    Joining_Date = table.Column<DateOnly>(type: "date", nullable: true),
                    Address = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Image_Id = table.Column<int>(type: "int", nullable: true),
                    Profile_ImageImage_Id = table.Column<int>(type: "int", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_EmployeesDB", x => x.Employee_Id);
                    table.ForeignKey(
                        name: "FK_EmployeesDB_DepartmentDB_Department_Id",
                        column: x => x.Department_Id,
                        principalTable: "DepartmentDB",
                        principalColumn: "Department_Id");
                    table.ForeignKey(
                        name: "FK_EmployeesDB_ImageDB_Profile_ImageImage_Id",
                        column: x => x.Profile_ImageImage_Id,
                        principalTable: "ImageDB",
                        principalColumn: "Image_Id");
                });

            migrationBuilder.CreateIndex(
                name: "IX_EmployeesDB_Department_Id",
                table: "EmployeesDB",
                column: "Department_Id");

            migrationBuilder.CreateIndex(
                name: "IX_EmployeesDB_Profile_ImageImage_Id",
                table: "EmployeesDB",
                column: "Profile_ImageImage_Id");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "EmployeesDB");

            migrationBuilder.DropTable(
                name: "DepartmentDB");

            migrationBuilder.DropTable(
                name: "ImageDB");
        }
    }
}
