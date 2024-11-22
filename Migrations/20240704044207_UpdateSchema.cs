using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

#pragma warning disable CA1814 // Prefer jagged arrays over multidimensional

namespace CrudUsingAjax.Migrations
{
    /// <inheritdoc />
    public partial class UpdateSchema : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.InsertData(
                table: "DepartmentDB",
                columns: new[] { "Department_Id", "DepartmentName" },
                values: new object[,]
                {
                    { 1, 1 },
                    { 2, 2 },
                    { 3, 3 },
                    { 4, 4 }
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DeleteData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 1);

            migrationBuilder.DeleteData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 2);

            migrationBuilder.DeleteData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 3);

            migrationBuilder.DeleteData(
                table: "DepartmentDB",
                keyColumn: "Department_Id",
                keyValue: 4);
        }
    }
}
