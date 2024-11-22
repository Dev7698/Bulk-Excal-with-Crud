using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace CrudUsingAjax.Migrations
{
    /// <inheritdoc />
    public partial class Employee : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id",
                table: "EmployeesDB");

            migrationBuilder.DropForeignKey(
                name: "FK_EmployeesDB_ImageDB_Profile_ImageImage_Id",
                table: "EmployeesDB");

            migrationBuilder.DropTable(
                name: "ImageDB");

            migrationBuilder.DropIndex(
                name: "IX_EmployeesDB_Department_Id",
                table: "EmployeesDB");

            migrationBuilder.DropColumn(
                name: "Image_Id",
                table: "EmployeesDB");

            migrationBuilder.RenameColumn(
                name: "Profile_ImageImage_Id",
                table: "EmployeesDB",
                newName: "Department_Id1");

            migrationBuilder.RenameIndex(
                name: "IX_EmployeesDB_Profile_ImageImage_Id",
                table: "EmployeesDB",
                newName: "IX_EmployeesDB_Department_Id1");

            migrationBuilder.AlterColumn<int>(
                name: "Department_Id",
                table: "EmployeesDB",
                type: "int",
                nullable: false,
                defaultValue: 0,
                oldClrType: typeof(int),
                oldType: "int",
                oldNullable: true);

            migrationBuilder.AddColumn<string>(
                name: "Profile_Image",
                table: "EmployeesDB",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id1",
                table: "EmployeesDB",
                column: "Department_Id1",
                principalTable: "DepartmentDB",
                principalColumn: "Department_Id");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id1",
                table: "EmployeesDB");

            migrationBuilder.DropColumn(
                name: "Profile_Image",
                table: "EmployeesDB");

            migrationBuilder.RenameColumn(
                name: "Department_Id1",
                table: "EmployeesDB",
                newName: "Profile_ImageImage_Id");

            migrationBuilder.RenameIndex(
                name: "IX_EmployeesDB_Department_Id1",
                table: "EmployeesDB",
                newName: "IX_EmployeesDB_Profile_ImageImage_Id");

            migrationBuilder.AlterColumn<int>(
                name: "Department_Id",
                table: "EmployeesDB",
                type: "int",
                nullable: true,
                oldClrType: typeof(int),
                oldType: "int");

            migrationBuilder.AddColumn<int>(
                name: "Image_Id",
                table: "EmployeesDB",
                type: "int",
                nullable: true);

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

            migrationBuilder.CreateIndex(
                name: "IX_EmployeesDB_Department_Id",
                table: "EmployeesDB",
                column: "Department_Id");

            migrationBuilder.AddForeignKey(
                name: "FK_EmployeesDB_DepartmentDB_Department_Id",
                table: "EmployeesDB",
                column: "Department_Id",
                principalTable: "DepartmentDB",
                principalColumn: "Department_Id");

            migrationBuilder.AddForeignKey(
                name: "FK_EmployeesDB_ImageDB_Profile_ImageImage_Id",
                table: "EmployeesDB",
                column: "Profile_ImageImage_Id",
                principalTable: "ImageDB",
                principalColumn: "Image_Id");
        }
    }
}
