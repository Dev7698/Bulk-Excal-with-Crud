using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace CrudUsingAjax.Migrations
{
    /// <inheritdoc />
    public partial class UpdateSchema1 : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.UpdateData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 1,
                column: "DepartmentName",
                value: 0);

            migrationBuilder.UpdateData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 2,
                column: "DepartmentName",
                value: 1);

            migrationBuilder.UpdateData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 3,
                column: "DepartmentName",
                value: 2);

            migrationBuilder.UpdateData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 4,
                column: "DepartmentName",
                value: 3);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.UpdateData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 1,
                column: "DepartmentName",
                value: 1);

            migrationBuilder.UpdateData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 2,
                column: "DepartmentName",
                value: 2);

            migrationBuilder.UpdateData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 3,
                column: "DepartmentName",
                value: 3);

            migrationBuilder.UpdateData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 4,
                column: "DepartmentName",
                value: 4);
        }
    }
}
