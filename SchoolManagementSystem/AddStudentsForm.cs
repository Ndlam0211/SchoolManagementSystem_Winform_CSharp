using System;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Drawing;
using System.Linq;
using ClosedXML.Excel;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Reflection;
using DocumentFormat.OpenXml.Office2010.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using DocumentFormat.OpenXml.Office.Word;

namespace SchoolManagementSystem
{
    public partial class AddStudentsForm : UserControl
    {
        private DashboardForm dashboardForm;
        public DashboardForm DashboardFormReference { get; set; }

        SqlConnection connectt = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;Initial Catalog=School;Integrated Security=True;");

        SqlConnection connect = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\ADMIN\OneDrive\Documents\school.mdf;Integrated Security=True;Connect Timeout=30");
        public AddStudentsForm(DashboardForm dashboardForm)
        {
            InitializeComponent();

            displayStudentData();

            this.dashboardForm = dashboardForm;
        }

        public void displayStudentData()
        {
            AddStudentsData adData = new AddStudentsData();

            student_gridData.DataSource = adData.studentData();
        }

        public void clearFields()
        {
            student_id.Text = "";
            student_name.Text = "";
            student_gender.SelectedIndex = -1;
            student_address.Text = "";
            student_grade.SelectedIndex = -1;
            student_section.SelectedIndex = -1;
            student_status.SelectedIndex = -1;
            student_image.Image = null;
            imagePath = "";
        }

        public string imagePath;
        private void student_browseBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Image files (*.bmp; *.jpg; *.jpeg; *.gif; *.png)|*.bmp;*.jpg;*.jpeg;*.gif;*.png";

            if (open.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Kiểm tra kích thước tệp
                    FileInfo fileInfo = new FileInfo(open.FileName);
                    long fileSizeInBytes = fileInfo.Length;
                    double fileSizeInMB = fileSizeInBytes / (1024.0 * 1024.0);

                    if (fileSizeInMB <= 3)
                    {
                        imagePath = open.FileName;

                        // Kiểm tra xem imagePath có giá trị không trước khi gán ImageLocation
                        if (!string.IsNullOrEmpty(imagePath))
                        {
                            student_image.ImageLocation = imagePath;
                        }
                        else
                        {
                            MessageBox.Show("Invalid image path.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("File size exceeds 3MB. Please choose a smaller file.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred.: {ex.Message}", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void student_updateBtn_Click(object sender, EventArgs e)
        {
            if (student_id.Text == ""
                || student_name.Text == ""
                || student_gender.Text == ""
                || student_address.Text == ""
                || student_grade.Text == ""
                || student_section.Text == ""
                || student_status.Text == ""
                || student_image.Image == null
                || imagePath == null)
            {
                MessageBox.Show("Please select item first", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (connect.State != ConnectionState.Open)
                {
                    try
                    {
                        connect.Open();

                        if (!IsValidInput(student_name.Text.Trim()))
                        {
                            MessageBox.Show("Please enter a valid full name.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        DialogResult check = MessageBox.Show("Are you sure you want to Update Student ID: "
                            + student_id.Text.Trim() + "?", "Confirmation Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        string checkStudentID = "SELECT COUNT(*) FROM students WHERE student_id = @studentID";

                        String updateData = "UPDATE students SET student_name = @studentName, " +
                                            "student_gender = @studentGender, student_adress = @studentAddress, " +
                                            "student_grade = @studentGrade, student_section = @studentSection, " +
                                            "student_status = @studentStatus, date_update = @dateUpdate " +
                                            "WHERE student_id = @studentID";

                        if (check == DialogResult.Yes)
                        {
                            using (SqlCommand checkSID = new SqlCommand(checkStudentID, connect))
                            {
                                checkSID.Parameters.AddWithValue("@studentID", student_id.Text.Trim());
                                int count = (int)checkSID.ExecuteScalar();

                                if (count >= 1)
                                {
                                    MessageBox.Show("Student ID: " + student_id.Text.Trim() + " is already exist"
                                        , "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }
                                else
                                {
                                    using(SqlCommand cmd = new SqlCommand(updateData, connect))
                                    {
                                        DateTime today = DateTime.Today;

                                        
                                        cmd.Parameters.AddWithValue("@studentName", student_name.Text.Trim());
                                        cmd.Parameters.AddWithValue("@studentGender", student_gender.Text.Trim());
                                        cmd.Parameters.AddWithValue("@studentAddress", student_address.Text.Trim());
                                        cmd.Parameters.AddWithValue("@studentGrade", student_grade.Text.Trim());
                                        cmd.Parameters.AddWithValue("@studentSection", student_section.Text.Trim());
                                        cmd.Parameters.AddWithValue("@studentStatus", student_status.Text.Trim());
                                        cmd.Parameters.AddWithValue("@dateUpdate", today.ToString());
                                        cmd.Parameters.AddWithValue("@studentID", student_id.Text.Trim());

                                        cmd.ExecuteNonQuery();
                                    }    
                                    displayStudentData();

                                    MessageBox.Show("Updated successfully!", "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    clearFields();
                                }

                               

                            }
                        }
                        else
                        {
                            MessageBox.Show("Cancelled.", "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            clearFields();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error connecting Database: " + ex, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                    finally
                    {
                        connect.Close();
                    }
                }
            }
        }

        private void student_gridData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                DataGridViewRow row = student_gridData.Rows[e.RowIndex];

                // Duyệt qua tất cả các ô trong hàng và đặt thuộc tính Selected
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.Selected = true;
                }

                student_id.Text = GetCellValueAsString(row, 1);
                student_name.Text = GetCellValueAsString(row, 2);
                student_gender.Text = GetCellValueAsString(row, 3);
                student_address.Text = GetCellValueAsString(row, 4);
                student_grade.Text = GetCellValueAsString(row, 5);
                student_section.Text = GetCellValueAsString(row, 6);

                imagePath = GetCellValueAsString(row, 7);
                LoadTeacherImage(imagePath);

                student_status.Text = GetCellValueAsString(row, 8);
            }
        }

        private string GetCellValueAsString(DataGridViewRow row, int cellIndex)
        {
            object cellValue = row.Cells[cellIndex].Value;
            return cellValue != null ? cellValue.ToString() : string.Empty;
        }

        private void LoadTeacherImage(string imagePath)
        {
            try
            {
                if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                {
                    student_image.Image = Image.FromFile(imagePath);
                }
                else
                {
                    student_image.Image = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while loading the image: {ex.Message}", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                student_image.Image = null;
            }
        }

        private void student_deleteBtn_Click(object sender, EventArgs e)
        {
            if (student_id.Text == "")
            {
                MessageBox.Show("Please select item first", "Error Message"
                    , MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (connect.State != ConnectionState.Open)
                {
                    DialogResult check = MessageBox.Show("Are you sure you want to Delete Student ID: "
                        + student_id.Text + "?", "Confirmation Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (check == DialogResult.Yes)
                    {

                        try
                        {
                            connect.Open();
                            DateTime today = DateTime.Today;

                            string deleteData = "UPDATE students SET date_delete = @dateDelete " +
                                "WHERE student_id = @studentID";

                            using (SqlCommand cmd = new SqlCommand(deleteData, connect))
                            {
                                cmd.Parameters.AddWithValue("@dateDelete", today.ToString());
                                cmd.Parameters.AddWithValue("@studentID", student_id.Text.Trim());

                                cmd.ExecuteNonQuery();

                                displayStudentData();

                                MessageBox.Show("Deleted successfully!", "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                clearFields();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error connecting Database: " + ex, "Error Message"
                        , MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            connect.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Cancelled.", "Information Message"
                        , MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
        }

        private void student_addBtn_Click(object sender, EventArgs e)
        {
            if (student_id.Text == ""
                || student_name.Text == ""
                || student_gender.Text == ""
                || student_address.Text == ""
                || student_grade.Text == ""
                || student_section.Text == ""
                || student_status.Text == ""
                || student_image.Image == null
                || imagePath == null)
            {
                MessageBox.Show("Please fill all blank fields", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (connect.State != ConnectionState.Open)
                {
                    try
                    {
                        connect.Open();

                        if (!IsValidInput(student_name.Text.Trim()))
                        {
                            MessageBox.Show("Please enter a valid full name.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        string checkStudentID = "SELECT COUNT(*) FROM students WHERE student_id = @studentID";

                        using (SqlCommand checkSID = new SqlCommand(checkStudentID, connect))
                        {
                            checkSID.Parameters.AddWithValue("@studentID", student_id.Text.Trim());
                            int count = (int)checkSID.ExecuteScalar();

                            if (count >= 1)
                            {
                                MessageBox.Show("Student ID: " + student_id.Text.Trim() + " is already exist"
                                    , "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            }
                            else
                            {
                                DateTime today = DateTime.Today;
                                string insertData = "INSERT INTO students (student_id, student_name" +
                                    ", student_gender, student_adress, student_grade, student_section" +
                                    ", student_image, student_status, date_insert) " +
                                    "VALUES(@studentID, @studentName, @studentGender, @studentAddress" +
                                    ", @studentGrade, @studentSection, @studentImage, @studentStatus" +
                                    ", @dateInsert)";

                                string path = Path.Combine(@"D:\HocTap\DesktopApp_C#_Winform\SchoolManagementSystem\SchoolManagementSystem\Student_Directory\", student_id.Text.Trim() + ".jpg");

                                string directoryPath = Path.GetDirectoryName(path);

                                if (!Directory.Exists(directoryPath))
                                {
                                    Directory.CreateDirectory(directoryPath);
                                }

                                File.Copy(imagePath, path, true);

                                using (SqlCommand cmd = new SqlCommand(insertData, connect))
                                {
                                    cmd.Parameters.AddWithValue("@studentID", student_id.Text.Trim());
                                    cmd.Parameters.AddWithValue("@studentName", student_name.Text.Trim());
                                    cmd.Parameters.AddWithValue("@studentGender", student_gender.Text.Trim());
                                    cmd.Parameters.AddWithValue("@studentAddress", student_address.Text.Trim());
                                    cmd.Parameters.AddWithValue("@studentGrade", student_grade.Text.Trim());
                                    cmd.Parameters.AddWithValue("@studentSection", student_section.Text.Trim());
                                    cmd.Parameters.AddWithValue("@studentImage", path);
                                    cmd.Parameters.AddWithValue("@studentStatus", student_status.Text.Trim());
                                    cmd.Parameters.AddWithValue("@dateInsert", today.ToString());

                                    cmd.ExecuteNonQuery();

                                    displayStudentData();

                                    MessageBox.Show("Added successfully!", "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    clearFields();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error connecting Database: " + ex, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                    finally
                    {
                        connect.Close();
                        dashboardForm.UpdateStudentData();
                    }
                }
            }
        }

        private void student_clearBtn_Click(object sender, EventArgs e)
        {
            clearFields();
        }

        private bool IsValidInput(string input)
        {
            return !string.IsNullOrEmpty(input) && input.All(c => char.IsLetter(c) || char.IsWhiteSpace(c) || char.GetUnicodeCategory(c) == System.Globalization.UnicodeCategory.LowercaseLetter || char.GetUnicodeCategory(c) == System.Globalization.UnicodeCategory.UppercaseLetter);
        }

        //private void student_importBtn_Click(object sender, EventArgs e)
        //{
        //    OpenFileDialog openFileDialog = new OpenFileDialog
        //    {
        //        Filter = "Excel Files|*.xls;*.xlsx"
        //    };

        //    if (openFileDialog.ShowDialog() == DialogResult.OK)
        //    {
        //        try
        //        {
        //            using (var workbook = new XLWorkbook(openFileDialog.FileName))
        //            {
        //                Kiểm tra định dạng file Excel
        //                if (!IsValidExcelFormat(workbook))
        //                {
        //                    return;
        //                }

        //                var worksheet = workbook.Worksheet(1);

        //                Assuming that your Excel columns are in the same order as your DataGridView columns
        //                DataTable dt = new DataTable();
        //                foreach (var firstRowCell in worksheet.FirstRow().Cells())
        //                {
        //                    dt.Columns.Add(firstRowCell.Value.ToString());
        //                }

        //                foreach (var row in worksheet.RowsUsed().Skip(1))
        //                {
        //                    dt.Rows.Add(row.Cells().Select(c => c.Value).ToArray());
        //                }

        //                Kiểm tra cấu trúc dữ liệu
        //                if (!IsDataTableStructureValid(dt))
        //                {
        //                    return;
        //                }

        //                Kiểm tra dữ liệu hợp lệ
        //                if (!IsImportDataValid(dt))
        //                {
        //                    return;
        //                }

        //                if (connect.State != ConnectionState.Open)
        //                {
        //                    try
        //                    {
        //                        connect.Open();

        //                        if (IsImportDataValid(dt))
        //                        {
        //                            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connect))
        //                            {
        //                                bulkCopy.DestinationTableName = "students";
        //                                bulkCopy.WriteToServer(dt);
        //                            }

        //                            displayStudentData();

        //                            MessageBox.Show("Import successful!", "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

        //                            LogImportActivity("Import successful", openFileDialog.FileName);
        //                        }
        //                        else
        //                        {
        //                            MessageBox.Show("Invalid data in Excel file. Please check and try again.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //                            LogImportActivity("Import failed - Invalid data", openFileDialog.FileName);
        //                        }
        //                    }
        //                    catch (SqlException ex)
        //                    {
        //                        MessageBox.Show($"Database Error: {ex.Message}", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                        LogImportActivity($"Error connecting to Database: {ex.Message}", openFileDialog.FileName);
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        MessageBox.Show($"Error importing data: {ex.Message}", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                        LogImportActivity($"Error importing data: {ex.Message}", openFileDialog.FileName);
        //                    }
        //                    finally
        //                    {
        //                        connect.Close();
        //                    }
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show($"Error importing data: {ex.Message}", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //            LogImportActivity($"Error importing data: {ex.Message}", openFileDialog.FileName);
        //        }
        //    }
        //}

        private void student_importBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                ImportExcelData(filePath);
            }
        }

        private void ImportExcelData(string filePath)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet == null)
                    {
                        MessageBox.Show("No worksheet found in the Excel file.");
                        return;
                    }

                    DataTable dt = new DataTable();

                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    // Add columns to DataTable
                    for (int col = worksheet.Dimension.Start.Column; col <= colCount; col++)
                    {
                        dt.Columns.Add(worksheet.Cells[1, col].Text);
                    }

                    // Add rows to DataTable
                    for (int row = worksheet.Dimension.Start.Row+1; row <= rowCount; row++)
                    {
                        List<string> listRows = new List<string>();
                        for (int col = worksheet.Dimension.Start.Column; col <= colCount; col++)
                        {
                            listRows.Add(worksheet.Cells[row, col].Text);
                        }
                        dt.Rows.Add(listRows.ToArray());
                    }

                    if (!IsDataTableStructureValid(dt))
                    {
                        return;
                    }

                    // Check for duplicate student_id
                    if (HasDuplicateStudentId(dt))
                    {
                        MessageBox.Show("Excel file contains duplicate 'student_id'. Import aborted.");
                        return;
                    }

                    // Save data to the database
                    using (SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\ADMIN\OneDrive\Documents\school.mdf;Integrated Security=True;Connect Timeout=30"))
                    {
                        connection.Open();
                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                        {
                            bulkCopy.DestinationTableName = "students";

                            bulkCopy.ColumnMappings.Add("ID", "id");
                            bulkCopy.ColumnMappings.Add("StudentID", "student_id");
                            bulkCopy.ColumnMappings.Add("StudentName", "student_name");
                            bulkCopy.ColumnMappings.Add("StudentGender", "student_gender");
                            bulkCopy.ColumnMappings.Add("StudentAddress", "student_adress");
                            bulkCopy.ColumnMappings.Add("StudentGrade", "student_grade");
                            bulkCopy.ColumnMappings.Add("StudentSection", "student_section");
                            bulkCopy.ColumnMappings.Add("StudentImage", "student_image");
                            bulkCopy.ColumnMappings.Add("Status", "student_status");
                            bulkCopy.ColumnMappings.Add("DateInsert", "date_insert");

                            bulkCopy.WriteToServer(dt);
                        }
                        connection.Close();
                    }

                    // Display data in DataGridView
                    displayStudentData();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error importing data: " + ex.Message);
            }
        }
        private bool HasDuplicateStudentId(DataTable dt)
        {
            using (SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\ADMIN\OneDrive\Documents\school.mdf;Integrated Security=True;Connect Timeout=30"))
            {
                connection.Open();

                // Lấy danh sách các student_id từ bảng students trong cơ sở dữ liệu
                SqlCommand cmd = new SqlCommand("SELECT student_id FROM students", connection);
                SqlDataReader reader = cmd.ExecuteReader();

                // Tạo một HashSet để lưu trữ các giá trị student_id từ cơ sở dữ liệu
                HashSet<string> dbStudentIds = new HashSet<string>();

                while (reader.Read())
                {
                    string dbStudentId = reader["student_id"].ToString();
                    dbStudentIds.Add(dbStudentId);
                }

                reader.Close();

                // Kiểm tra từng giá trị trong DataTable có tồn tại trong HashSet không
                foreach (DataRow row in dt.Rows)
                {
                    string dtStudentId = row["studentId"].ToString();

                    if (dbStudentIds.Contains(dtStudentId))
                    {
                        return true; // Có giá trị trùng lặp
                    }
                }

                return false; // Không có giá trị trùng lặp
            }
        }

        private bool IsDataTableStructureValid(DataTable dataTable)
        {
            if (dataTable.Columns.Count != 10)
            {
                MessageBox.Show("Invalid data structure. The Excel file must have 10 columns.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void student_exportBtn_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                FileName = "StudentData.xlsx"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("StudentData");

                        // Assuming that your DataGridView columns are in the same order as your Excel columns
                        foreach (DataGridViewColumn col in student_gridData.Columns)
                        {
                            worksheet.Cell(1, col.Index + 1).Value = col.HeaderText;
                        }

                        for (int i = 0; i < student_gridData.Rows.Count; i++)
                        {
                            for (int j = 0; j < student_gridData.Columns.Count; j++)
                            {
                                worksheet.Cell(i + 2, j + 1).Value = student_gridData.Rows[i].Cells[j].Value.ToString();
                            }
                        }

                        workbook.SaveAs(saveFileDialog.FileName);

                        MessageBox.Show("Export successful!", "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error exporting data: {ex.Message}", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

    }
}
