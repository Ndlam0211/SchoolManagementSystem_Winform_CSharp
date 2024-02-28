using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using ClosedXML.Excel;
using OfficeOpenXml;

namespace SchoolManagementSystem
{
    public partial class AddTeachersForm : UserControl
    {
        SqlConnection connect = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\ADMIN\OneDrive\Documents\school.mdf;Integrated Security=True;Connect Timeout=30");
        public AddTeachersForm()
        {
            InitializeComponent();

            teacherDisplayData();
        }

        public void teacherDisplayData()
        {
            AddTeachersData addTD = new AddTeachersData();

            teacher_gridData.DataSource = addTD.teacherData();
        }

        private void teacher_addBtn_Click(object sender, EventArgs e)
        {
            if (teacher_id.Text == ""
                || teacher_name.Text == ""
                || teacher_gender.Text == ""
                || teacher_address.Text == ""
                || teacher_status.Text == ""
                || teacher_image.Image == null
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

                        if (!IsValidInput(teacher_name.Text.Trim()))
                        {
                            MessageBox.Show("Please enter a valid full name.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        string checkTeacherID = "SELECT COUNT(*) FROM teachers WHERE teacher_id = @teacherID";

                        using (SqlCommand checkTID = new SqlCommand(checkTeacherID, connect))
                        {
                            checkTID.Parameters.AddWithValue("@teacherID", teacher_id.Text.Trim());
                            int count = (int)checkTID.ExecuteScalar();

                            if (count >= 1)
                            {
                                MessageBox.Show("Teacher ID: " + teacher_id.Text.Trim() + " is already exist"
                                    , "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            }
                            else
                            {
                                DateTime today = DateTime.Today;
                                string insertData = "INSERT INTO teachers " +
                                    "(teacher_id, teacher_name, teacher_gender, teacher_adress, " +
                                    "teacher_image, teacher_status, date_insert) " +
                                    "VALUES(@teacherID, @teacherName, @teacherGender, @teacherAddress," +
                                    "@teacherImage, @teacherStatus, @dateInsert)";

                                string path = Path.Combine(@"D:\HocTap\DesktopApp_C#_Winform\SchoolManagementSystem\SchoolManagementSystem\Teacher_Directory\", teacher_id.Text.Trim() + ".jpg");

                                string directoryPath = Path.GetDirectoryName(path);

                                if (!Directory.Exists(directoryPath))
                                {
                                    Directory.CreateDirectory(directoryPath);
                                }

                                File.Copy(imagePath, path, true);

                                using (SqlCommand cmd = new SqlCommand(insertData, connect))
                                {
                                    cmd.Parameters.AddWithValue("@teacherID", teacher_id.Text.Trim());
                                    cmd.Parameters.AddWithValue("@teacherName", teacher_name.Text.Trim());
                                    cmd.Parameters.AddWithValue("@teacherGender", teacher_gender.Text.Trim());
                                    cmd.Parameters.AddWithValue("@teacherAddress", teacher_address.Text.Trim());
                                    cmd.Parameters.AddWithValue("@teacherImage", path);
                                    cmd.Parameters.AddWithValue("@teacherStatus", teacher_status.Text.Trim());
                                    cmd.Parameters.AddWithValue("@dateInsert", today.ToString());

                                    cmd.ExecuteNonQuery();

                                    teacherDisplayData();

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
                    }
                }
            }
        }

        // Kiểm tra xem chuỗi không rỗng và không chứa ký tự đặc biệt nguyên tố
        private bool IsValidInput(string input)
        {
            return !string.IsNullOrEmpty(input) && input.All(c => char.IsLetter(c) || char.IsWhiteSpace(c) || char.GetUnicodeCategory(c) == System.Globalization.UnicodeCategory.LowercaseLetter || char.GetUnicodeCategory(c) == System.Globalization.UnicodeCategory.UppercaseLetter);
        }


        public void clearFields()
        {
            teacher_id.Text = "";
            teacher_name.Text = "";
            teacher_gender.SelectedIndex = -1;
            teacher_address.Text = "";
            teacher_status.SelectedIndex = -1;
            teacher_image.Image = null;
            imagePath = "";
        }

        private string imagePath;
        private void teacher_browseBtn_Click(object sender, EventArgs e)
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
                            teacher_image.ImageLocation = imagePath;
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

        private void teacher_clearBtn_Click(object sender, EventArgs e)
        {
            clearFields();
        }

        private void teacher_updateBtn_Click(object sender, EventArgs e)
        {
            if (teacher_id.Text == ""
                || teacher_name.Text == ""
                || teacher_gender.Text == ""
                || teacher_address.Text == ""
                || teacher_status.Text == ""
                || teacher_image.Image == null
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

                        if (!IsValidInput(teacher_name.Text.Trim()))
                        {
                            MessageBox.Show("Please enter a valid full name.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        DialogResult check = MessageBox.Show("Are you sure you want to Update Teacher ID: "
                            + teacher_id.Text.Trim() + "?", "Confirmation Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (check == DialogResult.Yes)
                        {
                            DateTime today = DateTime.Today;

                            String updateData = "UPDATE teachers SET " +
                                "teacher_name = @teacherName, teacher_gender = @teacherGender" +
                                ", teacher_adress = @teacherAddress" +
                                ", teacher_status = @teacherStatus" +
                                ", date_update = @dateUpdate WHERE teacher_id = @teacherID";


                            using (SqlCommand cmd = new SqlCommand(updateData, connect))
                            {
                                cmd.Parameters.AddWithValue("@teacherName", teacher_name.Text.Trim());
                                cmd.Parameters.AddWithValue("@teacherGender", teacher_gender.Text.Trim());
                                cmd.Parameters.AddWithValue("@teacherAddress", teacher_address.Text.Trim());
                                cmd.Parameters.AddWithValue("@teacherStatus", teacher_status.Text.Trim());
                                cmd.Parameters.AddWithValue("@dateUpdate", today.ToString());
                                cmd.Parameters.AddWithValue("@teacherID", teacher_id.Text.Trim());

                                cmd.ExecuteNonQuery();

                                teacherDisplayData();

                                MessageBox.Show("Updated successfully!", "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                clearFields();

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

        private void teacher_gridData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                DataGridViewRow row = teacher_gridData.Rows[e.RowIndex];

                // Duyệt qua tất cả các ô trong hàng và đặt thuộc tính Selected
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.Selected = true;
                }

                teacher_id.Text = GetCellValueAsString(row, 1);
                teacher_name.Text = GetCellValueAsString(row, 2);
                teacher_gender.Text = GetCellValueAsString(row, 3);
                teacher_address.Text = GetCellValueAsString(row, 4);

                imagePath = GetCellValueAsString(row, 5);
                LoadTeacherImage(imagePath);

                teacher_status.Text = GetCellValueAsString(row, 6);
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
                    teacher_image.Image = Image.FromFile(imagePath);
                }
                else
                {
                    teacher_image.Image = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while loading the image: {ex.Message}", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                teacher_image.Image = null;
            }
        }

        private void teacher_deleteBtn_Click(object sender, EventArgs e)
        {
            if (teacher_id.Text == "")
            {
                MessageBox.Show("Please select item first", "Error Message"
                    , MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (connect.State != ConnectionState.Open)
                {
                    DialogResult check = MessageBox.Show("Are you sure you want to Delete Teacher ID: "
                        + teacher_id.Text + "?", "Confirmation Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (check == DialogResult.Yes)
                    {

                        try
                        {
                            connect.Open();
                            DateTime today = DateTime.Today;

                            string deleteData = "UPDATE teachers SET date_delete = @dateDelete " +
                                "WHERE teacher_id = @teacherID";

                            using (SqlCommand cmd = new SqlCommand(deleteData, connect))
                            {
                                cmd.Parameters.AddWithValue("@dateDelete", today.ToString());
                                cmd.Parameters.AddWithValue("@teacherID", teacher_id.Text.Trim());

                                cmd.ExecuteNonQuery();

                                teacherDisplayData();

                                MessageBox.Show("Deleted successfully!", "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                clearFields();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error  connecting Database: " + ex, "Error Message"
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

        private void teacher_importBtn_Click(object sender, EventArgs e)
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
                    for (int row = worksheet.Dimension.Start.Row + 1; row <= rowCount; row++)
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
                        MessageBox.Show("Excel file contains duplicate 'teacher_id'. Import aborted.");
                        return;
                    }

                    // Save data to the database
                    using (SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\ADMIN\OneDrive\Documents\school.mdf;Integrated Security=True;Connect Timeout=30"))
                    {
                        connection.Open();
                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                        {
                            bulkCopy.DestinationTableName = "teachers";

                            bulkCopy.ColumnMappings.Add("ID", "id");
                            bulkCopy.ColumnMappings.Add("TeacherID", "teacher_id");
                            bulkCopy.ColumnMappings.Add("TeacherName", "teacher_name");
                            bulkCopy.ColumnMappings.Add("TeacherGender", "teacher_gender");
                            bulkCopy.ColumnMappings.Add("TeacherAdress", "teacher_adress");
                            bulkCopy.ColumnMappings.Add("StudentImage", "teacher_image");
                            bulkCopy.ColumnMappings.Add("Status", "teacher_status");
                            bulkCopy.ColumnMappings.Add("DateInsert", "date_insert");

                            bulkCopy.WriteToServer(dt);
                        }
                        connection.Close();
                    }

                    // Display data in DataGridView
                    teacherDisplayData();

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
                SqlCommand cmd = new SqlCommand("SELECT teacher_id FROM teachers", connection);
                SqlDataReader reader = cmd.ExecuteReader();

                // Tạo một HashSet để lưu trữ các giá trị student_id từ cơ sở dữ liệu
                HashSet<string> dbTeacherIds = new HashSet<string>();

                while (reader.Read())
                {
                    string dbTeacherId = reader["teacher_id"].ToString();
                    dbTeacherIds.Add(dbTeacherId);
                }

                reader.Close();

                // Kiểm tra từng giá trị trong DataTable có tồn tại trong HashSet không
                foreach (DataRow row in dt.Rows)
                {
                    string dtTeacherId = row["teacherId"].ToString();

                    if (dbTeacherIds.Contains(dtTeacherId))
                    {
                        return true; // Có giá trị trùng lặp
                    }
                }

                return false; // Không có giá trị trùng lặp
            }
        }

        private bool IsDataTableStructureValid(DataTable dataTable)
        {
            if (dataTable.Columns.Count != 8)
            {
                MessageBox.Show("Invalid data structure. The Excel file must have 10 columns.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void teacher_exportBtn_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                FileName = "TeacherData.xlsx"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("TeacherData");

                        // Assuming that your DataGridView columns are in the same order as your Excel columns
                        foreach (DataGridViewColumn col in teacher_gridData.Columns)
                        {
                            worksheet.Cell(1, col.Index + 1).Value = col.HeaderText;
                        }

                        for (int i = 0; i < teacher_gridData.Rows.Count; i++)
                        {
                            for (int j = 0; j < teacher_gridData.Columns.Count; j++)
                            {
                                worksheet.Cell(i + 2, j + 1).Value = teacher_gridData.Rows[i].Cells[j].Value.ToString();
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
