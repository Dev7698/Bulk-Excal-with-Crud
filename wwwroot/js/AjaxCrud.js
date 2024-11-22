
$(document).ready(function () {
    // Initialize Datatable 
    $("#EmployeeTable").DataTable({
        processing: true,
        serverSide: true,
        filter: true,
        orderMulti: true,
        Sorting: [[0, "desc"]],
        responsive: true,
        scrollCollapse: false,
        order: [],
        scrollX: true,
        ordering: true,
        lengthChange: true,
        paging: true,
        Paginate: true,
        pagingType: "full_numbers",
        pageLength: 10,
        "ajax": {
            "url": '/Employee/EmployeeList',
            "type": "GET",
            "dataType": "json",
            dataSrc: function (json) {
                //console.log(json);
                return json.data;
            }
        },
        "columns": [
            {
                "render": function (data, type, row) {
                    return `<div class="d-flex justify-content-between">
                                <input type="checkbox" class="select-checkbox" data-id="${row.employee_Id}">
                                <a class="btn btn-success edit-btn" data-id="${row.employee_Id}"><i class="bi bi-pencil-square"></i></a>
                                <a class="btn btn-danger delete-btn" data-id="${row.employee_Id}"><i class="bi bi-trash"></i></a>
                            </div>`;
                }
            },
            { "data": "first_Name", "name": "First_Name", "autoWidth": true },
            { "data": "last_Name", "name": "Last_Name", "autoWidth": true },
            { "data": "email", "name": "Email", "autoWidth": true },
            { "data": "phone_Number", "name": "Phone_Number", "autoWidth": true },
            { "data": "gender", "name": "Gender", "autoWidth": true },
            { "data": "departmentName", "name": "DepartmentName", "autoWidth": true },
            { "data": "joining_Date", "name": "Joining_Data", "autoWidth": true },
            { "data": "address", "name": "Address", "autoWidth": true },
        ]
    });


    // Open model for add new data
    $('#addEmpBtn').click(function () {
        $('#addEmpModel').modal('show');
    });

    // Open model for edit data
    $(document).on('click', '.edit-btn', function () {
        var id = $(this).data('id');
        console.log("==> Edit Id: ", id)
        LoadEmpData(id);
        $('#editEmpModel').modal('show')
    })

    // Open model for upload employee data
    $('#uploadEmpData').click(function () {
        $('#UploadFile').modal('show');
    });

    // UploadData
    $("#UploadEmpData").click(function () {
        var formData = new FormData();
        var file = $("#SelectedFile")[0].files[0];
        formData.append("file", file);

        $.ajax({
            type: "POST",
            url: "/Employee/UploadEmpData",
            data: formData,
            processData: false,
            contentType: false,
            xhrFields: {
                responseType: 'blob'  // to handle file download
            },
            success: function (response, status, xhr) {
                // Close the upload modal
                $("#UploadFile").modal("hide");

                // Reload the data table  
                $("#EmployeeTable").DataTable().ajax.reload();

                var disposition = xhr.getResponseHeader('Content-Disposition');

                if (disposition && disposition.indexOf('attachment') !== -1) {
                    // Handle file download
                    var filename = "";
                    var matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition);
                    if (matches != null && matches[1]) filename = matches[1].replace(/['"]/g, '');

                    var type = xhr.getResponseHeader('Content-Type');
                    var blob = new Blob([response], { type: type });

                    if (typeof window.navigator.msSaveBlob !== 'undefined') {
                        // IE workaround
                        window.navigator.msSaveBlob(blob, filename);
                    } else {
                        var URL = window.URL || window.webkitURL;
                        var downloadUrl = URL.createObjectURL(blob);

                        if (filename) {
                            var a = document.createElement("a");
                            if (typeof a.download === 'undefined') {
                                window.location = downloadUrl;
                            } else {
                                a.href = downloadUrl;
                                a.download = filename;
                                document.body.appendChild(a);
                                a.click();
                            }
                        } else {
                            window.location = downloadUrl;
                        }
                        setTimeout(function () { URL.revokeObjectURL(downloadUrl); }, 100); // cleanup
                    }
                } else {
                    // Handle JSON response
                    var reader = new FileReader();
                    reader.onload = function (event) {
                        var jsonResponse = JSON.parse(event.target.result);
                        if (jsonResponse.success) {
                            $("#succesModel").removeClass('hide').addClass('show');
                            $('#succesModel').text(jsonResponse.message);
                        } else {

                            $('#errorModal').removeClass('hide').addClass('show');
                            $('#errorModal').text(jsonResponse.error);
                        }
                    };
                    reader.readAsText(response);
                }
            },

            error: function (error) {
                console.log(error);
                $("#UploadFile").modal("hide");
                $('#errorModal').removeClass('hide').addClass('show');

                $('#errorModal').text("Failed Upload");

            }
        });
    });


    //$("#UploadEmpData").click(function () {

    //    var formData = new FormData();
    //    var file = $("#SelectedFile")[0].files[0];
    //    formData.append("File", file);

    //    //console.log("==>>", formData);

    //    $.ajax({
    //        type: "POST",
    //        url: "/Employee/UploadEmpData",
    //        data: formData,
    //        processData: false,
    //        contentType: false,
    //        success: function (response) {
    //            if (response.success) {
    //                $("#UploadFile").modal("hide");
    //                $("#EmployeeTable").DataTable().ajax.reload();
    //                $("#succesModel").removeClass('hide').addClass('show');
    //                $('#succesModel').text(response.message);
    //            }
    //            else {
    //                $('#UploadFile').modal('hide');
    //                $('#errorModal').removeClass('hide').addClass('show');
    //                $('#errorModal').text(response.error);
    //            }
    //        },
    //        error: function (error) {
    //            console.log(error);
    //        }
    //    })
    //})

    // Download Data
    //$("#downloadExcel").on('click', function () {
    //    window.location.href = "/Employee/DownloadEmpExcel"
    //});
    $("#downloadExcel").on('click', function () {
        let selectedIds = [];
        $('.select-checkbox:checked').each(function () {
            selectedIds.push($(this).data('id'));
        });

        if (selectedIds.length > 0) {
            $.ajax({
                url: "/Employee/DownloadEmpExcel",
                type: "POST",

                contentType: "application/json",
                data: JSON.stringify(selectedIds),
                success: function (response) {
                    if (response.success) {
                        // Redirect to the download URL
                        window.location.href = response.downloadUrl;
                    } else {
                        alert(response.message); // Show error message
                    }
                },
                error: function () {
                    alert('Error downloading file.');
                }
            });
        } else {
            alert('Please select at least one employee.');
        }
    });


    // Save Data
    $("#SaveEmployeeData").click(function () {
        var fname = $("#fname").val();
        var lname = $("#lname").val();
        var email = $("#email").val();
        var phone = $("#phone").val();
        var gender = $("#gender").val();
        var department = $("#department").val();
        var jdate = $("#jdate").val();
        var address = $("#address").val();

        var fileInput = document.getElementById('image');
        var image = fileInput.files[0];

        //console .log("==> Department", department);     
        //console.log("==> Image :- ", image);            

        if ($("#employeeForm").valid()) {
            var formData = new FormData();
            formData.append("First_Name", fname);
            formData.append("Last_Name", lname);
            formData.append("Email", email);
            formData.append("Phone_Number", phone);
            formData.append("Gender", gender);
            formData.append("Department_Id", department);
            formData.append("Joining_Date", jdate);
            formData.append("Address", address);
            if (image) {
                formData.append("image", image);
            }

            $.ajax({
                type: "POST",
                url: "/Employee/AddNewEmpData",
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    if (response.success) {
                        // console.log("Add Data ==>", response);
                        $('#employeeForm')[0].reset();
                        $('#addEmpModel').modal('hide');
                        $("#EmployeeTable").DataTable().ajax.reload();
                        $("#succesModel").removeClass('hide').addClass('show');
                        $("#succesModel").text(response.message);
                    }
                    else {
                        $('#addEmpModel').modal('hide');
                        $('#errorModal').removeClass('hide').addClass('show');
                        $('#errorModal').text(response.error);

                    }
                },
                error: function (error) {
                    console.log(error);
                }
            })
        }
        else {
            $("#employeeForm").validate().focusInvalid();
        }
    })

    // Update Data
    $("#UpdateEmployeeData").click(function () {

        if ($("#edit_employeeForm").valid()) {
            var formData = new FormData($("#edit_employeeForm")[0]);
            formData.append("Employee_Id", $("#edit_employee_id").val());
            $.ajax({
                type: "POST",
                url: "/Employee/UpdateEmpData",
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    // console.log("==>>", response);
                    if (response.success) {
                        $('#employeeForm')[0].reset();
                        $("#editEmpModel").modal('hide');
                        $("#EmployeeTable").DataTable().ajax.reload();
                        $("#succesModel").removeClass("hide").addClass("show");
                        $("#succesModel").text(response.message);
                    }
                    else {
                        $('#editEmpModel').modal('hide');
                        $('#errorModal').removeClass('hide').addClass('show');
                        $('#errorModal').text(response.error);
                    }
                },
                error: function (error) {
                    console.log(error)
                }
            })
        }
        else {
            $("#edit_employeeForm").validate().focusInvalid();
        }
    });

    // Delete
    $(document).on('click', '.delete-btn', function () {
        var id = $(this).data('id');
        DeleteEmployeeData(id);
    });

    // Load employee data
    function LoadEmpData(id) {
        $.ajax({
            type: 'GET',
            url: '/Employee/GetEmpData',
            data: { id: id },
            success: function (response) {
                if (response) {
                    // console.log("==> Load : ", response)
                    var employee = response.data;
                    $("#edit_employee_id").val(employee?.employee_Id);
                    $("#edit_fname").val(employee?.first_Name);
                    $("#edit_lname").val(employee?.last_Name);
                    $("#edit_email").val(employee?.email);
                    $("#edit_phone").val(employee?.phone_Number);
                    if (employee?.gender === 'Male') {
                        $("#edit_gender_male").prop("checked", true);
                    } else if (employee?.gender === "Female") {
                        $("#edit_gender_Female").prop("checked", true);
                    }
                    $("#edit_department").val(employee?.department_Id);
                    $("#edit_jdate").val(employee?.joining_Date);
                    $("#edit_address").val(employee?.address);
                } else {
                    console.log(response.message)
                }
            },
            error: function (error) {
                console.log(error)
            }
        });
    }

    // Delete Data
    function DeleteEmployeeData(id) {
        $.ajax({
            type: 'POST',
            url: '/Employee/DeleteEmpData',
            data: { id: id },
            success: function (response) {
                if (response.success) {
                    $("#EmployeeTable").DataTable().ajax.reload();
                    $("#succesModel").removeClass('hide alert-success').addClass('show alert-danger');
                    $("#succesModel").text(response.message);
                } else {
                    console.error("Failed to delete employee.");
                    $('#errorModal').removeClass('hide').addClass('show');
                    $('#errorModal').text(response.error);
                }
            },
            error: function (error) {
                console.error(error);
            }
        });
    }

    // Close model
    $('#closeModel').on('click', function () {
        $('#employeeForm')[0].reset();
        $('#employeeModal').hide();
    });
    $("#Upload_Data").validate({
        errorClass: 'text-danger',
        rules: {
            SelectedFile: {
                required: true,
            }
        },
        messages: {
            SelectedFile: {
                required: "Please select a file."
            }
        },
        submitHandler: function (form) {
            form.submit();
        }
    });

    $("#employeeForm").validate({
        errorClass: 'text-danger',
        rules: {
            fname: {
                required: true,
                minlength: 2
            },
            lname: {
                required: true,
                minlength: 2
            },
            email: {
                required: true,
                email: true
            },
            phone: {
                required: true,
                digits: true,
                minlength: 10,
                maxlength: 10
            },
            department: {
                required: true
            },
            jdate: {
                required: true,
                date: true
            },
            address: {
                required: true,
            }
        },
        messages: {
            fname: {
                required: "Please enter your first name.",
                minlength: "Your first name must be at least 2 characters long."
            },
            lname: {
                required: "Please enter your last name.",
                minlength: "Your last name must be at least 2 characters long."
            },

            email: {
                required: "Please enter your email address.",
                email: "Please enter a valid email address."
            },
            phone: {
                required: "Please enter your phone number.",
                digits: "Please enter only digits.",
                minlength: "Please enter a valid 10-digit phone number.",
                maxlength: "Please enter a valid 10-digit phone number."
            },
            department: {
                required: "Please select a department."
            },
            jdate: {
                required: "Please enter a joining date.",
                date: "Please enter a valid date."
            },
            address: {
                required: "Please enter your address."
            }
        },
        submitHandler: function (form) {
            form.submit();
        }
    });

    $("#edit_employeeForm").validate({
        errorClass: 'text-danger',
        rules: {
            edit_fname: {
                required: true,
                minlength: 4
            },
            edit_lname: {
                required: true,
                minlength: 4
            },
            edit_email: {
                required: true,
                email: true
            },
            edit_phone: {
                required: true,
                digits: true,
                minlength: 10,
                maxlength: 10
            },
            edit_department: {
                required: true
            },
            edit_jdate: {
                required: true,
                date: true
            },
            edit_address: {
                required: true,
                // Remove digits rule to allow for alphabets in address
            }
        },
        messages: {
            edit_fname: {
                required: "Please enter your first name.",
                minlength: "Your first name must be at least 4 characters long."
            },
            edit_lname: {
                required: "Please enter your last name.",
                minlength: "Your last name must be at least 4 characters long."
            },
            edit_email: {
                required: "Please enter your email address.",
                email: "Please enter a valid email address."
            },
            edit_phone: {
                required: "Please enter your phone number.",
                digits: "Please enter only digits.",
                minlength: "Please enter a valid 10-digit phone number.",
                maxlength: "Please enter a valid 10-digit phone number."
            },
            edit_department: {
                required: "Please select a department."
            },
            edit_jdate: {
                required: "Please enter a joining date.",
                date: "Please enter a valid date."
            },
            edit_address: {
                required: "Please enter your address."
                // Removed digits validation, allowing any characters
            }
        },
        submitHandler: function (form) {
            form.submit();
        }
    });

    // Form validation
    //$("#Upload_Data").validate({
    //    errorClass: 'text-danger',
    //    rules: {
    //        SelectedFile: {
    //            required: true,
    //        }
    //    },
    //    messages: {
    //        SelectedFile: {
    //            required: "Please select the file"
    //        }
    //    },
    //    submitHandler: function (form) {
    //        form.submit();
    //    }
    //});
    //$("#employeeForm").validate({
    //    errorClass: 'text-danger',
    //    rules: {
    //        fname: {
    //            required: true,
    //            minlength: 2
    //        },
    //        lname: {
    //            required: true,
    //            minlength: 2
    //        },
    //        email: {
    //            required: true,
    //            email: true
    //        },
    //        phone: {
    //            required: true,
    //            digits: true,
    //            minlength: 10,
    //            maxlength: 10,
    //        },
    //        department: {
    //            required: true
    //        },
    //        jdate: {
    //            required: true,
    //            date: true
    //        },
    //        address: {
    //            required: true,
    //        }
    //    },
    //    messages: {
    //        fname: {
    //            required: "Please enter your first name",
    //            minlength: "Your first name must be at least 2 characters long"
    //        },
    //        lname: {
    //            required: "Please enter your last name",
    //            minlength: "Your last name must be at least 2 characters long"
    //        },
    //        email: {
    //            required: "Please enter your email address",
    //            email: "Please enter a valid email address"
    //        },
    //        phone: {
    //            required: "Please enter your phone number",
    //            digits: "Please enter only digits",
    //            minlength: "Please enter a valid mobile number",
    //            maxlength: "Please enter a valid mobile number"
    //        },
    //        department: {
    //            required: "Please select a department"
    //        },
    //        jdate: {
    //            required: "Please enter a joining date",
    //            date: "Please enter a valid date"
    //        },
    //        address: {
    //            required: "Please enter your address"
    //        }
    //    },
    //    submitHandler: function (form) {
    //        form.submit();
    //    }
    //});

    //$("#edit_employeeForm").validate({
    //    errorClass: 'text-danger',
    //    rules: {
    //        edit_fname: {
    //            required: true,
    //            minlength: 4
    //        },
    //        edit_lname: {
    //            required: true,
    //            minlength: 4
    //        },
    //        edit_email: {
    //            required: true,
    //            email: true
    //        },
    //        edit_phone: {
    //            required: true,
    //            digits: true,
    //            minlength: 10,
    //            maxlength: 10
    //        },
    //        edit_department: {
    //            required: true
    //        },
    //        edit_jdate: {
    //            required: true,
    //            date: true
    //        },
    //        edit_address: {
    //            digits: false
    //            required: true,
    //        }
    //    },
    //    messages: {
    //        eidt_fname: {
    //            required: "Please enter your first name",
    //            minlength: "Your first name must be at least 2 characters long"
    //        },
    //        edit_lname: {
    //            required: "Please enter your last name",
    //            minlength: "Your last name must be at least 2 characters long"
    //        },
    //        edit_email: {
    //            required: "Please enter your email address",
    //            email: "Please enter a valid email address"
    //        },
    //        edit_phone: {
    //            required: "Please enter your phone number",
    //            digits: "Please enter only digits",
    //            minlength: "Please enter a valid mobile number",
    //            maxlength: "Please enter a valid mobile number"
    //        },
    //        edit_department: {
    //            required: "Please select a department"
    //        },
    //        edit_jdate: {
    //            required: "Please enter a joining date",
    //            date: "Please enter a valid date"
    //        },
    //        edit_address: {
    //            required: "Please enter your address"
    //             digits: "please enter only Alphabets"
    //        }
    //    },
    //    submitHandler: function (form) {
    //        form.submit();
    //    }
    //});

});
