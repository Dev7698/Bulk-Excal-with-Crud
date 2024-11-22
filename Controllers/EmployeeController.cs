using CrudUsingAjax.DataBase;
using CrudUsingAjax.Models;
using CrudUsingAjax.Repositories;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Elfie.Diagnostics;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Mail;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace CrudUsingAjax.Controllers
{
    public class EmployeeController : Controller
    {
        private readonly IEmployeeRepository<Employee> _employeeRepository;
        private readonly IEmployeeRepository<Department> _departmentRepository;
        private readonly ILogger<EmployeeController> _logger;
        private readonly IConfiguration _configuration;

        public EmployeeController(IEmployeeRepository<Employee> employeeRepository, IEmployeeRepository<Department> departmentRepository, ILogger<EmployeeController> logger)
        {
            _employeeRepository = employeeRepository;
            _departmentRepository = departmentRepository;
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }


        public JsonResult EmployeeList() 
        {
            try
            {
                // Retrieve query parameters
                var draw = HttpContext.Request.Query["draw"].FirstOrDefault();
                var start = HttpContext.Request.Query["start"].FirstOrDefault();
                var length = HttpContext.Request.Query["length"].FirstOrDefault();
                var searchValue = HttpContext.Request.Query["search[Value]"].FirstOrDefault();
                var sortColumn = HttpContext.Request.Query["order[0][column]"].FirstOrDefault();
                var sortColumnDirection = HttpContext.Request.Query["order[0][dir]"].FirstOrDefault();

                int skip = start != null ? int.Parse(start) : 0;
                int pageLength = length != null ? int.Parse(length) : 0;

                var employeeData = _employeeRepository.GetAll().AsQueryable();

                // Apply search filter
                if (!string.IsNullOrWhiteSpace(searchValue))
                {
                    employeeData = employeeData.Where(m => m.First_Name.Contains(searchValue) ||
                                                            m.Last_Name.Contains(searchValue) ||
                                                            m.Email.Contains(searchValue) ||
                                                            m.Phone_Number.Contains(searchValue) ||
                                                            m.Gender.Contains(searchValue) ||
                                                            m.Address.Contains(searchValue));
                }

                // Apply sorting
                employeeData = sortColumn switch
                {
                    "1" => sortColumnDirection == "asc" ? employeeData.OrderBy(e => e.First_Name) : employeeData.OrderByDescending(e => e.First_Name),
                    "2" => sortColumnDirection == "asc" ? employeeData.OrderBy(e => e.Last_Name) : employeeData.OrderByDescending(e => e.Last_Name),
                    "3" => sortColumnDirection == "asc" ? employeeData.OrderBy(e => e.Email) : employeeData.OrderByDescending(e => e.Email),
                    "4" => sortColumnDirection == "asc" ? employeeData.OrderBy(e => e.Phone_Number) : employeeData.OrderByDescending(e => e.Phone_Number),
                    "5" => sortColumnDirection == "asc" ? employeeData.OrderBy(e => e.Gender) : employeeData.OrderByDescending(e => e.Gender),
                    "6" => sortColumnDirection == "asc" ? employeeData.OrderBy(e => e.Department.DepartmentName) : employeeData.OrderByDescending(e => e.Department.DepartmentName),
                    "7" => sortColumnDirection == "asc" ? employeeData.OrderBy(e => e.Joining_Date) : employeeData.OrderByDescending(e => e.Joining_Date),
                    "8" => sortColumnDirection == "asc" ? employeeData.OrderBy(e => e.Address) : employeeData.OrderByDescending(e => e.Address),
                    _ => sortColumnDirection == "asc" ? employeeData.OrderBy(e => e.First_Name) : employeeData.OrderByDescending(e => e.First_Name),
                };

                int recordsTotal = employeeData.Count();
                var data = employeeData.Skip(skip).Take(pageLength).Select(e => new
                {
                    employee_Id = e.Employee_Id,
                    first_Name = e.First_Name,
                    last_Name = e.Last_Name,
                    email = e.Email,
                    phone_Number = e.Phone_Number,
                    gender = e.Gender,
                    departmentName = e.Department.DepartmentName.ToString(),
                    joining_Date = e.Joining_Date,
                    address = e.Address,
                }).ToList();

                var JsonData = new { draw = Convert.ToInt32(draw), recordsFiltered = recordsTotal, recordsTotal = recordsTotal, data = data };
                return Json(JsonData);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error  fetching employee list");
                return Json(new { success = false });
            }
        }

        [HttpGet]


        public JsonResult GetEmpData(int id)
        {
            try
            {
                var employee = _employeeRepository.GetById(id);
                if (employee != null)
                {
                    return Json(new
                    {
                        success = true,
                        data = new
                        {
                            employee_Id = employee.Employee_Id,
                            first_Name = employee.First_Name,
                            last_Name = employee.Last_Name,
                            email = employee.Email,
                            phone_Number = employee.Phone_Number,
                            gender = employee.Gender,
                            department_Id = employee.Department_Id,
                            joining_Date = employee.Joining_Date,
                            address = employee.Address,
                        }
                    });
                }
                return Json(new { success = false, error = "Employee not found" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred while fetching employee data");
                return Json(new { success = false, data = ex.Message });
            }
        }

        [HttpPost]
        public async Task<JsonResult> AddNewEmpData([FromForm] Employee employeeData, [FromForm] IFormFile image)
        {
            try
            {
                // Check for existing employee by email
                var existingEmployee = _employeeRepository.GetAll().FirstOrDefault(e => e.Email == employeeData.Email);
                if (existingEmployee != null)
                {
                    return Json(new
                    {
                        success = false,
                        error = "Email already exists."
                    });
                }


                if (image != null)
                {
                    var uniqueFileName = $"{Guid.NewGuid()}_{Path.GetFileName(image.FileName)}";
                    var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/image", uniqueFileName);
                    using (var stream = new FileStream(imagePath, FileMode.Create))
                    {
                        image.CopyTo(stream);
                    }
                    employeeData.Profile_Image = "/image/" + uniqueFileName;
                }

                var department = _departmentRepository.GetById(employeeData.Department_Id);
                if (department == null)
                {
                    _logger.LogError("Invalid Department_Id: " + employeeData.Department_Id);
                    return Json(new
                    {
                        success = false,
                        error = "Cancel form submission"
                    });
                }

                // Save new employee data
                _employeeRepository.Add(employeeData);
                _employeeRepository.SaveChanges();

                return Json(new { success = true, data = employeeData, message = "Employee  successfully  added " });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return Json(new
                {
                    success = false,
                });
            }
        }


        public JsonResult UpdateEmpData([FromForm] Employee employee, [FromForm] IFormFile image)
        {
            try
            {
                var existedEmployeeData = GetExistingEmployee(employee.Employee_Id);
                if (existedEmployeeData == null)
                {
                    return GenerateErrorResponse("Employee not found");
                }

                if (IsEmailInUse(employee.Email, existedEmployeeData.Email))
                {
                    return GenerateErrorResponse("Email already exists.");
                }

                UpdateEmployeeFields(employee, existedEmployeeData);

                _employeeRepository.Update(employee);
                _employeeRepository.SaveChanges();

                return Json(new
                {
                    success = true,
                    data = employee,
                    message = "Employee updated  successfully "
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return GenerateErrorResponse("An error occurred while updating the employee.");
            }
        }

        private Employee GetExistingEmployee(int employeeId)
        {
            return _employeeRepository.GetAll().AsNoTracking().FirstOrDefault(e => e.Employee_Id == employeeId);
        }

        private bool IsEmailInUse(string newEmail, string existingEmail)
        {
            if (!string.IsNullOrWhiteSpace(newEmail) && newEmail != existingEmail)
            {
                return _employeeRepository.GetAll().Any(e => e.Email == newEmail);
            }
            return false;
        }

        private void UpdateEmployeeFields(Employee employee, Employee existedEmployeeData)
        {
            employee.First_Name = string.IsNullOrWhiteSpace(employee.First_Name) ? existedEmployeeData.First_Name : employee.First_Name;
            employee.Last_Name = string.IsNullOrWhiteSpace(employee.Last_Name) ? existedEmployeeData.Last_Name : employee.Last_Name;
            employee.Email = string.IsNullOrWhiteSpace(employee.Email) ? existedEmployeeData.Email : employee.Email;
            employee.Phone_Number = string.IsNullOrWhiteSpace(employee.Phone_Number) ? existedEmployeeData.Phone_Number : employee.Phone_Number;
            employee.Gender = string.IsNullOrWhiteSpace(employee.Gender) ? existedEmployeeData.Gender : employee.Gender;
            employee.Department_Id = employee.Department_Id == 0 ? existedEmployeeData.Department_Id : employee.Department_Id;
            employee.Joining_Date = employee.Joining_Date == null ? existedEmployeeData.Joining_Date : employee.Joining_Date;
            employee.Address = string.IsNullOrWhiteSpace(employee.Address) ? existedEmployeeData.Address : employee.Address;
        }

        private JsonResult GenerateErrorResponse(string message)
        {
            return Json(new
            {
                success = true,
                error = message
            });
        }

        [HttpPost]


        public JsonResult DeleteEmpData(int id)
        {
            if (id <= 0)
            {
                return Json(new
                {
                    success = false,
                    message = "Invalid employee ID."
                });
            }

            try
            {
                var employee = _employeeRepository.GetById(id);
                if (employee == null)
                {
                    return Json(new
                    {
                        success = false,
                        message = "Employee not found."
                    });
                }

                _employeeRepository.Delete(id);
                _employeeRepository.SaveChanges();

                return Json(new
                {
                    success = true,
                    message = "Employee deleted successfully."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting employee with ID: {Id}", id);
                return Json(new
                {
                    success = false,
                    message = "An error occurred while deleting the employee."
                });
            }
        }

        //public IActionResult UploadEmpData(IFormFile file)
        //{
        //    var uploadResults = new List<(string Email, string Status, string Message)>();

        //    try
        //    {
        //        if (file == null || file.Length == 0)
        //        {
        //             return Json(new { success = false, error = "File is empty or not selected" });
        //        }

        //        string dataFileName = Path.GetFileName(file.FileName);
        //        string extension = Path.GetExtension(dataFileName).ToLower();
        //        var allowedExtensions = new[] { ".xls", ".xlsx", ".csv" };

        //        if (!allowedExtensions.Contains(extension))
        //        {
        //            return Json(new { success = false, error = "Only Excel files (.xls, .xlsx, .csv) are allowed." });
        //        }

        //        var uploadDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploaded_Files");

        //        if (!Directory.Exists(uploadDirectory))
        //        {
        //            Directory.CreateDirectory(uploadDirectory);
        //        }

        //        var uniqueFileName = $"{Guid.NewGuid()}_{Path.GetFileName(file.FileName)}";
        //        var filePath = Path.Combine(uploadDirectory, uniqueFileName);

        //        using (var stream = new FileStream(filePath, FileMode.Create))
        //        {
        //            file.CopyTo(stream);
        //        }

        //        var employeeList = new List<Employee>();

        //        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        //        {                                                                                       
        //            IExcelDataReader reader = null;

        //            if (extension == ".xls")
        //            {
        //                reader = ExcelReaderFactory.CreateBinaryReader(stream);
        //            }
        //            else if (extension == ".xlsx")
        //            {
        //                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        //            }
        //            else if (extension == ".csv")
        //            {
        //                reader = ExcelReaderFactory.CreateCsvReader(stream);
        //            }

        //            if (reader != null)
        //            {
        //                using (reader)
        //                {
        //                    while (reader.Read())
        //                    {
        //                        if (reader.Depth == 0) continue;

        //                        var firstName = reader.GetValue(1)?.ToString();
        //                        var lastName = reader.GetValue(2)?.ToString();
        //                        var email = reader.GetValue(3)?.ToString();
        //                        var phoneNumber = reader.GetValue(4)?.ToString();
        //                        var gender = reader.GetValue(5)?.ToString();
        //                        var departmentStr = reader.GetValue(6)?.ToString();
        //                        var joiningDateStr = reader.GetValue(7)?.ToString();
        //                        var address = reader.GetValue(8)?.ToString();
        //                        var imageFileName = reader.GetValue(9)?.ToString();

        //                        var errors = new List<string>();

        //                        if (string.IsNullOrWhiteSpace(firstName)) errors.Add("First Name is required");
        //                        if (string.IsNullOrWhiteSpace(lastName)) errors.Add("Last Name is required");
        //                        if (string.IsNullOrWhiteSpace(email)) errors.Add("Email is required");
        //                        if (string.IsNullOrWhiteSpace(phoneNumber)) errors.Add("Phone Number is required");
        //                        if (string.IsNullOrWhiteSpace(gender)) errors.Add("Gender is required");
        //                        if (string.IsNullOrWhiteSpace(departmentStr)) errors.Add("Department ID is required");
        //                        if (string.IsNullOrWhiteSpace(joiningDateStr)) errors.Add("Joining Date is required");
        //                        if (string.IsNullOrWhiteSpace(address)) errors.Add("Address is required");

        //                        if (errors.Count > 0)
        //                        {
        //                            uploadResults.Add((email, "Error", string.Join(" , ", errors)));
        //                            continue;
        //                        }

        //                        var departmentId = int.TryParse(departmentStr, out var deptId) ? deptId : 0;

        //                        if (_departmentRepository.GetById(departmentId) == null)
        //                        {
        //                            uploadResults.Add((email, "Error", $"Department ID : {departmentId} not found in database"));
        //                            continue;
        //                        }

        //                        var existingEmployee = _employeeRepository.GetAll().FirstOrDefault(e => e.Email == email);

        //                        if (existingEmployee != null)
        //                        {
        //                            uploadResults.Add((email, "Error", "Email already exists"));
        //                            continue;
        //                        }

        //                        var employee = new Employee
        //                        {
        //                            First_Name = firstName,
        //                            Last_Name = lastName,
        //                            Email = email,
        //                            Phone_Number = phoneNumber,
        //                            Gender = gender,
        //                            Department_Id = departmentId,
        //                            Joining_Date = DateTime.TryParse(joiningDateStr, out var joiningDate) ? new DateOnly?(DateOnly.FromDateTime(joiningDate)) : null,
        //                            Address = address,
        //                        };

        //                        if (!string.IsNullOrWhiteSpace(imageFileName))
        //                        {
        //                            var uniqueImageName = $"{Guid.NewGuid()}_{Path.GetFileName(imageFileName)}";
        //                            var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/image", uniqueImageName);

        //                            using (var imageStream = new FileStream(imagePath, FileMode.Create))
        //                            { 
        //                                file.CopyTo(imageStream);
        //                            }
        //                            employee.Profile_Image = "/image/" + uniqueImageName;
        //                        }
        //                        employeeList.Add(employee);
        //                    }
        //                }
        //            }
        //        } 

        //        if (employeeList.Count > 0)
        //        { 
        //            _employeeRepository.AddRange(employeeList);
        //            _employeeRepository.SaveChanges(); 
        //            foreach (var emp in employeeList) 
        //            {
        //                uploadResults.Add((emp.Email, "Success", "Employee successfully uploaded"));
        //            }
        //        }
        //        else
        //        {
        //            uploadResults.Add(("N/A", "Error", "No valid employee to upload."));
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex.Message);
        //        return Json(new { success = false, error = "Employee data not uploaded" });
        //    }

        //    // Generate response Excel 
        //    try
        //    {
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            using (var package = new ExcelPackage(memoryStream))
        //            {
        //                var worksheet = package.Workbook.Worksheets.Add("Upload Results");
        //                worksheet.Cells[1, 1].Value = "Email";
        //                worksheet.Cells[1, 2].Value = "Status";
        //                worksheet.Cells[1, 3].Value = "Message";

        //                for (int i = 0; i < uploadResults.Count; i++)
        //                {
        //                    worksheet.Cells[i + 2, 1].Value = uploadResults[i].Email;
        //                    worksheet.Cells[i + 2, 2].Value = uploadResults[i].Status;
        //                    worksheet.Cells[i + 2, 3].Value = uploadResults[i].Message;
        //                }

        //                package.Save(); 
        //            }

        //            var fileBytes = memoryStream.ToArray();
        //            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "UploadResults.xlsx"); 
        //        }     
        //    }     
        //    catch (Exception ex) 
        //    {
        //        _logger.LogError(ex.Message);
        //        return Json(new { success = false, error = "Failed to generate Excel file." }); 
        //    }
        //}

        //public IActionResult UploadEmpData(IFormFile file)
        //{
        //    var uploadResults = new List<(string Email, string Status, string Message)>();
        //    var errors = new List<(string Email, string FullRecord, string Message)>();

        //    try
        //    {
        //        if (file == null || file.Length == 0)
        //        {
        //            return Json(new { success = false, error = "File is empty or not selected" });
        //        }

        //        string dataFileName = Path.GetFileName(file.FileName);
        //        string extension = Path.GetExtension(dataFileName).ToLower();
        //        var allowedExtensions = new[] { ".xls", ".xlsx", ".csv" };

        //        if (!allowedExtensions.Contains(extension))
        //        {
        //            return Json(new { success = false, error = "Only Excel files (.xls, .xlsx, .csv) are allowed." });
        //        } 

        //        var uploadDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploaded_Files");

        //        if (!Directory.Exists(uploadDirectory))
        //        {
        //            Directory.CreateDirectory(uploadDirectory);
        //        }

        //        var uniqueFileName = $"{Guid.NewGuid()}_{Path.GetFileName(file.FileName)}";
        //        var filePath = Path.Combine(uploadDirectory, uniqueFileName);

        //        using (var stream = new FileStream(filePath, FileMode.Create))
        //        {
        //            file.CopyTo(stream);
        //        }

        //        var employeeList = new List<Employee>();

        //        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        //        {
        //            IExcelDataReader reader = extension switch
        //            {
        //                ".xls" => ExcelReaderFactory.CreateBinaryReader(stream),
        //                ".xlsx" => ExcelReaderFactory.CreateOpenXmlReader(stream),
        //                ".csv" => ExcelReaderFactory.CreateCsvReader(stream),
        //                _ => null
        //            };

        //            if (reader != null)
        //            {
        //                using (reader)
        //                {
        //                    while (reader.Read())
        //                    {
        //                        if (reader.Depth == 0) continue; // Skip header row

        //                        var firstName = reader.GetValue(1)?.ToString();
        //                        var lastName = reader.GetValue(2)?.ToString();
        //                        var email = reader.GetValue(3)?.ToString();
        //                        var phoneNumber = reader.GetValue(4)?.ToString();
        //                        var gender = reader.GetValue(5)?.ToString();
        //                        var departmentStr = reader.GetValue(6)?.ToString();
        //                        var joiningDateStr = reader.GetValue(7)?.ToString();
        //                        var address = reader.GetValue(8)?.ToString();
        //                        var imageFileName = reader.GetValue(9)?.ToString();

        //                        var fullRecord = $"{firstName}, {lastName}, {email}, {phoneNumber}, {gender}, {departmentStr}, {joiningDateStr}, {address}, {imageFileName}";

        //                        // Validate fields
        //                        var errorMessage = ValidateEmployeeData(email, firstName, lastName, phoneNumber, gender, departmentStr, joiningDateStr, address);
        //                        if (!string.IsNullOrWhiteSpace(errorMessage))
        //                        {
        //                            errors.Add((email, fullRecord, errorMessage));
        //                            continue; // Skip this record if there's an error
        //                        }

        //                        var departmentId = int.TryParse(departmentStr, out var deptId) ? deptId : 0;

        //                        if (_departmentRepository.GetById(departmentId) == null)
        //                        {
        //                            errors.Add((email, fullRecord, $"Department ID: {departmentId} not found in database"));
        //                            continue;
        //                        }

        //                        var existingEmployee = _employeeRepository.GetAll().FirstOrDefault(e => e.Email == email);

        //                        if (existingEmployee != null)
        //                        {
        //                            errors.Add((email, fullRecord, "Email already exists"));
        //                            continue;
        //                        }

        //                        var employee = new Employee
        //                        {
        //                            First_Name = firstName,
        //                            Last_Name = lastName,
        //                            Email = email,
        //                            Phone_Number = phoneNumber,
        //                            Gender = gender,
        //                            Department_Id = departmentId,
        //                            Joining_Date = DateTime.TryParse(joiningDateStr, out var joiningDate) ? new DateOnly?(DateOnly.FromDateTime(joiningDate)) : null,
        //                            Address = address,
        //                        };

        //                        if (!string.IsNullOrWhiteSpace(imageFileName))
        //                        {
        //                            var uniqueImageName = $"{Guid.NewGuid()}_{Path.GetFileName(imageFileName)}";
        //                            var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/image", uniqueImageName);

        //                            using (var imageStream = new FileStream(imagePath, FileMode.Create))
        //                            {
        //                                file.CopyTo(imageStream);
        //                            }
        //                            employee.Profile_Image = "/image/" + uniqueImageName;
        //                        }

        //                        employeeList.Add(employee);
        //                    }
        //                }
        //            }
        //        }

        //        if (employeeList.Count > 0)
        //        {
        //            _employeeRepository.AddRange(employeeList);
        //            _employeeRepository.SaveChanges();
        //            foreach (var emp in employeeList)
        //            {
        //                uploadResults.Add((emp.Email, "Success", "Employee successfully uploaded"));
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex.Message);
        //        return Json(new { success = false, error = "Employee data not uploaded" });
        //    }

        //    // Generate the appropriate response
        //    if (errors.Any())
        //    {
        //        return GenerateErrorReport(errors);
        //    }
        //    else
        //    {
        //        return GenerateUploadResultsReport(uploadResults);
        //    }
        //}

        //private string ValidateEmployeeData(string email, string firstName, string lastName, string phoneNumber, string gender, string departmentStr, string joiningDateStr, string address)
        //{
        //    var errors = new List<string>();

        //    if (string.IsNullOrWhiteSpace(firstName))
        //        errors.Add("First Name is required.");
        //    if (string.IsNullOrWhiteSpace(lastName))
        //        errors.Add("Last Name is required.");
        //    if (string.IsNullOrWhiteSpace(email))
        //        errors.Add("Email is required.");
        //    if (string.IsNullOrWhiteSpace(phoneNumber))
        //        errors.Add("Phone Number is required.");
        //    if (string.IsNullOrWhiteSpace(gender))
        //        errors.Add("Gender is required.");
        //    if (string.IsNullOrWhiteSpace(departmentStr))
        //        errors.Add("Department ID is required.");
        //    if (string.IsNullOrWhiteSpace(joiningDateStr))
        //        errors.Add("Joining Date is required.");
        //    if (string.IsNullOrWhiteSpace(address))
        //        errors.Add("Address is required.");

        //    return errors.Any() ? string.Join(" ", errors) : null;
        //}

        //private IActionResult GenerateErrorReport(List<(string Email, string FullRecord, string Message)> errors)
        //{
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //    using (var memoryStream = new MemoryStream())
        //    {
        //        using (var package = new ExcelPackage(memoryStream))
        //        {
        //            var worksheet = package.Workbook.Worksheets.Add("Error Report");
        //            worksheet.Cells[1, 1].Value = "Email";
        //            worksheet.Cells[1, 2].Value = "Full Record";
        //            worksheet.Cells[1, 3].Value = "Error Message";

        //            for (int i = 0; i < errors.Count; i++)
        //            {
        //                worksheet.Cells[i + 2, 1].Value = errors[i].Email;
        //                worksheet.Cells[i + 2, 2].Value = errors[i].FullRecord;
        //                worksheet.Cells[i + 2, 3].Value = errors[i].Message;
        //            }

        //            // Formatting the worksheet
        //            worksheet.Cells[1, 1, errors.Count + 1, 3].AutoFitColumns();
        //            worksheet.Cells[1, 1, 1, 3].Style.Font.Bold = true;

        //            package.Save();
        //        }

        //        var fileBytes = memoryStream.ToArray();
        //        return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ErrorReport.xlsx");
        //    }
        //}

        //private IActionResult GenerateUploadResultsReport(List<(string Email, string Status, string Message)> uploadResults)
        //{
        //    using (var memoryStream = new MemoryStream())
        //    {
        //        using (var package = new ExcelPackage(memoryStream))
        //        {
        //            var worksheet = package.Workbook.Worksheets.Add("Upload Results");
        //            worksheet.Cells[1, 1].Value = "Email";
        //            worksheet.Cells[1, 2].Value = "Status";
        //            worksheet.Cells[1, 3].Value = "Message";

        //            for (int i = 0; i < uploadResults.Count; i++)
        //            {
        //                worksheet.Cells[i + 2, 1].Value = uploadResults[i].Email;
        //                worksheet.Cells[i + 2, 2].Value = uploadResults[i].Status;
        //                worksheet.Cells[i + 2, 3].Value = uploadResults[i].Message;
        //            }

        //            // Formatting the worksheet
        //            worksheet.Cells[1, 1, uploadResults.Count + 1, 3].AutoFitColumns();
        //            worksheet.Cells[1, 1, 1, 3].Style.Font.Bold = true;

        //            package.Save();
        //        }

        //        var fileBytes = memoryStream.ToArray();
        //        return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "UploadResults.xlsx");
        //    }
        //}


        [HttpPost]
        public IActionResult UploadEmpData(IFormFile file)
        {

            var uploadResults = new List<(string Email, string FirstName, string LastName, string Status, string Message)>();
            var errors = new List<(string Email, string FirstName, string LastName, string FullRecord, string Message)>();

            try
            {
                if (file == null || file.Length == 0)
                {
                    return Json(new { success = false, error = "File is empty or not selected" });
                }

                string dataFileName = Path.GetFileName(file.FileName);
                string extension = Path.GetExtension(dataFileName).ToLower();
                var allowedExtensions = new[] { ".xls", ".xlsx", ".csv" };

                if (!allowedExtensions.Contains(extension))
                {
                    return Json(new { success = false, error = "Only Excel files (.xls, .xlsx, .csv) are allowed." });
                }

                var uploadDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploaded_Files");

                if (!Directory.Exists(uploadDirectory))
                {
                    Directory.CreateDirectory(uploadDirectory);
                }

                var uniqueFileName = $"{Guid.NewGuid()}_{Path.GetFileName(file.FileName)}";
                var filePath = Path.Combine(uploadDirectory, uniqueFileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                var employeeList = new List<Employee>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader = extension switch
                    {
                        ".xls" => ExcelReaderFactory.CreateBinaryReader(stream),
                        ".xlsx" => ExcelReaderFactory.CreateOpenXmlReader(stream),
                        ".csv" => ExcelReaderFactory.CreateCsvReader(stream),
                        _ => null
                    };

                    if (reader != null)
                    {
                        using (reader)
                        {
                            while (reader.Read())
                            {
                                if (reader.Depth == 0) continue; // Skip header row

                                var firstName = reader.GetValue(1)?.ToString();
                                var lastName = reader.GetValue(2)?.ToString();
                                var email = reader.GetValue(3)?.ToString();
                                var phoneNumber = reader.GetValue(4)?.ToString();
                                var gender = reader.GetValue(5)?.ToString();
                                var departmentStr = reader.GetValue(6)?.ToString();
                                var joiningDateStr = reader.GetValue(7)?.ToString();
                                var address = reader.GetValue(8)?.ToString();
                                var imageFileName = reader.GetValue(9)?.ToString();

                                var fullRecord = $"{firstName}, {lastName}, {email}, {phoneNumber}, {gender}, {departmentStr}, {joiningDateStr}, {address}, {imageFileName}";

                                // Validate fields
                                var errorMessage = ValidateEmployeeData(email, firstName, lastName, phoneNumber, gender, departmentStr, joiningDateStr, address);
                                if (!string.IsNullOrWhiteSpace(errorMessage))
                                {
                                    errors.Add((email, firstName, lastName, fullRecord, errorMessage));
                                    uploadResults.Add((email, firstName, lastName, "Failed", errorMessage));
                                    continue; // Skip this record if there's an error
                                }
                                 
                                var departmentId = int.TryParse(departmentStr, out var deptId) ? deptId : 0;

                                if (_departmentRepository.GetById(departmentId) == null)
                                {
                                    errors.Add((email, firstName, lastName, fullRecord, $"Department ID: {departmentId} not found in database"));
                                    uploadResults.Add((email, firstName, lastName, "Failed", $"Department ID: {departmentId} not found in database"));
                                    continue;
                                }

                                var existingEmployee = _employeeRepository.GetAll().FirstOrDefault(e => e.Email == email);

                                if (existingEmployee != null)
                                {
                                    errors.Add((email, firstName, lastName, fullRecord, "Email already exists"));
                                    uploadResults.Add((email, firstName, lastName, "Failed", "Email already exists"));
                                    continue;
                                }

                                var employee = new Employee
                                {
                                    First_Name = firstName,
                                    Last_Name = lastName,
                                    Email = email,
                                    Phone_Number = phoneNumber,
                                    Gender = gender,
                                    Department_Id = departmentId,
                                    Joining_Date = DateTime.TryParse(joiningDateStr, out var joiningDate) ? new DateOnly?(DateOnly.FromDateTime(joiningDate)) : null,
                                    Address = address,
                                };

                                if (!string.IsNullOrWhiteSpace(imageFileName))
                                {
                                    var uniqueImageName = $"{Guid.NewGuid()}_{Path.GetFileName(imageFileName)}";
                                    var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/image", uniqueImageName);

                                    using (var imageStream = new FileStream(imagePath, FileMode.Create))
                                    {
                                        file.CopyTo(imageStream);
                                    }
                                    employee.Profile_Image = "/image/" + uniqueImageName;
                                }

                                employeeList.Add(employee);
                                uploadResults.Add((email, firstName, lastName, "Success", "Employee successfully uploaded"));
                            }
                        }
                    }
                }

                if (employeeList.Count > 0)
                {
                    _employeeRepository.AddRange(employeeList);
                    _employeeRepository.SaveChanges();
                }

            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return Json(new { success = false, error = "Employee data not uploaded" });
            }

            // Generate a combined report for both successful and failed uploads
            return GenerateCombinedReport(uploadResults);
        }

        private IActionResult GenerateCombinedReport(List<(string Email, string FirstName, string LastName, string Status, string Message)> uploadResults)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var memoryStream = new MemoryStream())
            {
                using (var package = new ExcelPackage(memoryStream))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Upload Results");
                    worksheet.Cells[1, 1].Value = "Email";
                    worksheet.Cells[1, 2].Value = "First Name";
                    worksheet.Cells[1, 3].Value = "Last Name";
                    worksheet.Cells[1, 4].Value = "Status";
                    worksheet.Cells[1, 5].Value = "Message";


                    for (int i = 0; i < uploadResults.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = uploadResults[i].Email;
                        worksheet.Cells[i + 2, 2].Value = uploadResults[i].FirstName;
                        worksheet.Cells[i + 2, 3].Value = uploadResults[i].LastName;
                        worksheet.Cells[i + 2, 4].Value = uploadResults[i].Status;
                        worksheet.Cells[i + 2, 5].Value = uploadResults[i].Message;
                    }

                    // Formatting the worksheet
                    worksheet.Cells[1, 1, uploadResults.Count + 1, 5].AutoFitColumns();
                    worksheet.Cells[1, 1, 1, 5].Style.Font.Bold = true;

                    package.Save();
                }

                var fileBytes = memoryStream.ToArray();
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "UploadResult.xlsx");
            }
        }

        private string ValidateEmployeeData(string email, string firstName, string lastName, string phoneNumber, string gender, string departmentStr, string joiningDateStr, string address)
        {
            var errors = new List<string>();

            if (string.IsNullOrWhiteSpace(firstName))
                errors.Add("First Name is required.");
            if (string.IsNullOrWhiteSpace(lastName))
                errors.Add("Last Name is required.");
            if (string.IsNullOrWhiteSpace(email))
                errors.Add("Email is required.");
            if (string.IsNullOrWhiteSpace(phoneNumber))
                errors.Add("Phone Number is required.");
            if (string.IsNullOrWhiteSpace(gender))
                errors.Add("Gender is required.");
            if (string.IsNullOrWhiteSpace(departmentStr))
                errors.Add("Department ID is required.");
            if (string.IsNullOrWhiteSpace(joiningDateStr))
                errors.Add("Joining Date is required.");
            if (string.IsNullOrWhiteSpace(address))
                errors.Add("Address is required.");

            return errors.Any() ? string.Join(" ", errors) : null;
        }

        [HttpPost]
        public IActionResult DownloadEmpExcel([FromBody] List<int> selectedIds)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;

                // Fetch only the selected employees with their departments
                var employees = _employeeRepository.GetAll()
                    .Include(e => e.Department)
                    .Where(e => selectedIds.Contains(e.Employee_Id)) // Filter by selected IDs
                    .ToList();

                // Check if there are any selected employees
                if (!employees.Any())
                {
                    return Json(new { success = false, message = "No employee data available." });
                }

                // Create a DataTable for the employee data
                var datatable = new DataTable("Employees");
                datatable.Columns.AddRange(new DataColumn[]
                {
            new DataColumn("First Name"),
            new DataColumn("Last Name"),
            new DataColumn("Email"),
            new DataColumn("Phone Number"),
            new DataColumn("Gender"),
            new DataColumn("Department"),
            new DataColumn("Joining Date"),
            new DataColumn("Address"),
            new DataColumn("Profile Image")
                });

                // Populate the DataTable with selected employee data
                foreach (var employee in employees)
                {
                    datatable.Rows.Add(
                        employee.First_Name,
                        employee.Last_Name,
                        employee.Email,
                        employee.Phone_Number,
                        employee.Gender,
                        employee.Department?.DepartmentName,
                        employee.Joining_Date,
                        employee.Address,
                        employee.Profile_Image
                    );
                }

                // Create the Excel file in memory
                var memoryStream = new MemoryStream();
                using (var excelPackage = new ExcelPackage(memoryStream))
                {
                    var worksheet = excelPackage.Workbook.Worksheets.Add("Employees");
                    worksheet.Cells["A1"].LoadFromDataTable(datatable, true);

                    // Format the columns as needed
                    var totalRows = datatable.Rows.Count + 1; // Include header row
                    var textColumns = new[] { "B", "C", "E", "F", "H", "I", "J", "K" };
                    foreach (var column in textColumns)
                    {
                        worksheet.Cells[$"{column}2:{column}{totalRows}"].Style.Numberformat.Format = "@"; // Text format
                    }
                    worksheet.Cells["D2:D" + totalRows].Style.Numberformat.Format = "yyyy-mm-dd"; // Phone Number format
                    worksheet.Cells["G2:G" + totalRows].Style.Numberformat.Format = "@"; // Address format

                    // Save the file to memory
                    excelPackage.Save();
                    memoryStream.Position = 0;

                    // Save the file to a temporary location
                    var filePath = Path.Combine(Path.GetTempPath(), "Employees.xlsx");
                    System.IO.File.WriteAllBytes(filePath, memoryStream.ToArray());

                    // Return the download URL
                    return Json(new { success = true, downloadUrl = Url.Action("DownloadFile", new { filename = "Employees.xlsx" }) });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error downloading employee Excel file.");
                return Json(new { success = false, message = ex.Message });
            }
        }
        [HttpGet]
        public IActionResult DownloadFile(string filename)
        {
            var filePath = Path.Combine(Path.GetTempPath(), filename);
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound();
            }

            // Read file bytes
            var fileBytes = System.IO.File.ReadAllBytes(filePath);

            // Send email with the file attached
            SendEmailWithAttachment(filename, fileBytes);

            // Return file for download
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        }

        private void SendEmailWithAttachment(string filename, byte[] fileBytes)
        {
            string fromMail = "princep@keyconcepts.co.in";
            string fromPassword = "mysurat123#";
            string toMail = "dev.visitorz@gmail.com";

            using (var message = new MailMessage())
            {
                message.From = new MailAddress(fromMail);
                message.Subject = "Here is your Excel file";
                message.To.Add(new MailAddress(toMail));
                message.Body = "<html><body>Attached is your requested Excel file.</body></html>";
                message.IsBodyHtml = true;

                // Attach the file
                using (var stream = new MemoryStream(fileBytes))
                {
                    message.Attachments.Add(new Attachment(stream, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

                    using (var smtpClient = new SmtpClient("smtp.gmail.com"))
                    {
                        smtpClient.Port = 587;
                        smtpClient.Credentials = new NetworkCredential(fromMail, fromPassword);
                        smtpClient.EnableSsl = true;

                        smtpClient.Send(message);
                    }
                }
            }
        }



        //public IActionResult DownloadEmpExcel([FromBody] List<int> selectedIds)
        //{
        //    try
        //    {
        //        ExcelPackage.LicenseContext = LicenseContext.Commercial;

        //        // Fetch all employees with their departments
        //        var employees = _employeeRepository.GetAll().Include(e => e.Department).ToList();

        //        // Check if there are any employees
        //        if (!employees.Any())
        //        {
        //            return Json(new { success = false, message = "No employee data available." });
        //        }

        //        // Create a DataTable for the employee data
        //        var datatable = new DataTable("Employees");
        //        datatable.Columns.AddRange(new DataColumn[]
        //        {
        //    new DataColumn("First Name"),
        //    new DataColumn("Last Name"),
        //    new DataColumn("Email"),
        //    new DataColumn("Phone Number"),
        //    new DataColumn("Gender"),
        //    new DataColumn("Department"),
        //    new DataColumn("Joining Date"),
        //    new DataColumn("Address"),
        //    new DataColumn("Profile Image")// Add any additional fields you have
        //          // Example additional field
        //                                       // Add more columns as necessary
        //        });

        //        // Populate the DataTable with employee data
        //        foreach (var employee in employees)
        //        {
        //            datatable.Rows.Add(
        //                employee.First_Name,
        //                employee.Last_Name,
        //                employee.Email,
        //                employee.Phone_Number,
        //                employee.Gender,
        //                employee.Department?.DepartmentName,
        //                employee.Joining_Date, // Format the date
        //                employee.Address,
        //                employee.Profile_Image// Include profile image if necessary
        //                                      //  employee.Salary,        // Example additional field
        //                                      //   employee.Position       // Example additional field
        //                                      // Add more fields as necessary
        //            );
        //        }

        //        // Create the Excel file in memory
        //        var memoryStream = new MemoryStream();  
        //        using (var excelPackage = new ExcelPackage(memoryStream))
        //        {

        //            var worksheet = excelPackage.Workbook.Worksheets.Add("Employees");

        //            worksheet.Cells["A1"].LoadFromDataTable(datatable, true);

        //            // Format the columns as needed
        //            var totalRows = datatable.Rows.Count + 1; // Include header row
        //            var textColumns = new[] { "B", "C", "E", "F", "H", "I", "J", "K" }; // Adjust this for all text columns
        //            foreach (var column in textColumns)
        //            {
        //                 worksheet.Cells[$"{column}2:{column}{totalRows}"].Style.Numberformat.Format = "@"; // Text format
        //            }
        //            worksheet.Cells["D2:D" + totalRows].Style.Numberformat.Format = "yyyy-mm-dd"; // Adjust for Phone Number
        //            worksheet.Cells["G2:G" + totalRows].Style.Numberformat.Format = "@"; // Address format

        //            // Save and prepare the file for download
        //            excelPackage.Save();
        //                memoryStream.Position = 0;
        //                return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Employees.xlsx");
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex, "Error downloading employee Excel file.");
        //        return Json(new { success = false, message = ex.Message });
        //    }
        //}








        //public IActionResult DownloadEmpExcel()
        //{
        //    try
        //    {
        //        var employees = _employeeRepository.GetAll().Include(e => e.Department).ToList();

        //        if (!employees.Any())
        //        {
        //            return Json(new { success = false, message = "No employee data available." });
        //        }

        //        var datatable = new DataTable("Employees");
        //        datatable.Columns.AddRange(new DataColumn[]
        //        {
        //    new DataColumn("First Name"),
        //    new DataColumn("Last Name"),
        //    new DataColumn("Email"),
        //    new DataColumn("Phone Number"),
        //    new DataColumn("Gender"),
        //    new DataColumn("Department"),
        //    new DataColumn("Joining Date"),
        //    new DataColumn("Address"),
        //    new DataColumn("Profile Image")
        //        });

        //        foreach (var employee in employees)
        //        {
        //            datatable.Rows.Add(
        //                employee.First_Name,
        //                employee.Last_Name,
        //                employee.Email,
        //                employee.Phone_Number,
        //                employee.Gender,
        //                employee.Department?.DepartmentName,
        //                employee.Joining_Date,
        //                employee.Address,
        //               employee.Profile_Image // Ensure you add this if needed
        //            );
        //        }

        //        using (var memoryStream = new MemoryStream())
        //        using (var excelPackage = new ExcelPackage(memoryStream))
        //        { 
        //            var worksheet = excelPackage.Workbook.Worksheets.Add("Employees");
        //            worksheet.Cells["A1"].LoadFromDataTable(datatable, true);

        //            var totalRows = employees.Count + 1;
        //            var textColumns = new[] { "B", "C", "E", "F", "H", "I" };
        //            foreach (var column in textColumns)
        //            {
        //                worksheet.Cells[$"{column}2:{column}{totalRows}"].Style.Numberformat.Format = "@";
        //            }
        //            worksheet.Cells["D2:D" + totalRows].Style.Numberformat.Format = "yyyy-mm-dd"; // Adjusted for the correct column
        //            worksheet.Cells["G2:G" + totalRows].Style.Numberformat.Format = "@"; // Text format for Address

        //            excelPackage.Save();
        //            memoryStream.Position = 0;
        //            return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Employees.xlsx");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex, "Error downloading employee Excel file.");
        //        return Json(new { success = false, message = ex.Message });
        //    }
        //}




    }
}

