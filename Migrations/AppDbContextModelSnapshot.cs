﻿// <auto-generated />
using System;
using CrudUsingAjax.DataBase;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

#nullable disable

namespace CrudUsingAjax.Migrations
{
    [DbContext(typeof(AppDbContext))]
    partial class AppDbContextModelSnapshot : ModelSnapshot
    {
        protected override void BuildModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "8.0.6")
                .HasAnnotation("Relational:MaxIdentifierLength", 128);

            SqlServerModelBuilderExtensions.UseIdentityColumns(modelBuilder);

            modelBuilder.Entity("CrudUsingAjax.Models.Department", b =>
                {
                    b.Property<int>("Department_Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("Department_Id"));

                    b.Property<int?>("DepartmentName")
                        .HasColumnType("int");

                    b.HasKey("Department_Id");

                    b.ToTable("DepartmentDB");

                    b.HasData(
                        new
                        {
                            Department_Id = 1,
                            DepartmentName = 1
                        },
                        new
                        {
                            Department_Id = 2,
                            DepartmentName = 2
                        },
                        new
                        {
                            Department_Id = 3,
                            DepartmentName = 3
                        },
                        new
                        {
                            Department_Id = 4,
                            DepartmentName = 4
                        });
                });

            modelBuilder.Entity("CrudUsingAjax.Models.Employee", b =>
                {
                    b.Property<int>("Employee_Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("Employee_Id"));

                    b.Property<string>("Address")
                        .HasColumnType("nvarchar(max)");

                    b.Property<int>("Department_Id")
                        .HasColumnType("int");

                    b.Property<string>("Email")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("First_Name")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Gender")
                        .HasColumnType("nvarchar(max)");

                    b.Property<DateOnly?>("Joining_Date")
                        .HasColumnType("date");

                    b.Property<string>("Last_Name")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Phone_Number")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Profile_Image")
                        .HasColumnType("nvarchar(max)");

                    b.HasKey("Employee_Id");

                    b.HasIndex("Department_Id");

                    b.ToTable("EmployeesDB");
                });

            modelBuilder.Entity("CrudUsingAjax.Models.Employee", b =>
                {
                    b.HasOne("CrudUsingAjax.Models.Department", "Department")
                        .WithMany()
                        .HasForeignKey("Department_Id")
                        .OnDelete(DeleteBehavior.Restrict)
                        .IsRequired();

                    b.Navigation("Department");
                });
#pragma warning restore 612, 618
        }
    }
}
