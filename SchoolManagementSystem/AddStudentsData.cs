using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace SchoolManagementSystem
{
    class AddStudentsData
    {
        SqlConnection connect = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\ADMIN\OneDrive\Documents\school.mdf;Integrated Security=True;Connect Timeout=30");

        public int ID { set; get; }
        public string StudentID { set; get; }
        public string StudentName { set; get; }
        public string StudentGender { set; get; }
        public string StudentAddress { set; get; }
        public string StudentGrade { set; get; }
        public string StudentSection { set; get; }
        public string StudentImage { set; get; }
        public string Status { set; get; }
        public string DateInsert { set; get; }

        public List<AddStudentsData> studentData()
        {
            List<AddStudentsData> listData = new List<AddStudentsData>();
            if (connect.State != ConnectionState.Open)
            {
                try
                {
                    connect.Open();

                    string sql = "SELECT * FROM students WHERE date_delete IS NULL";

                    using (SqlCommand cmd = new SqlCommand(sql, connect))
                    {
                        SqlDataReader reader = cmd.ExecuteReader();

                        while (reader.Read())
                        {
                            AddStudentsData addSD = new AddStudentsData();
                            addSD.ID = (int)reader["id"];
                            addSD.StudentID = reader["student_id"].ToString();
                            addSD.StudentName = reader["student_name"].ToString();
                            addSD.StudentGender = reader["student_gender"].ToString();
                            addSD.StudentAddress = reader["student_adress"].ToString();
                            addSD.StudentGrade = reader["student_grade"].ToString();
                            addSD.StudentSection = reader["student_section"].ToString();
                            addSD.StudentImage = reader["student_image"].ToString();
                            addSD.Status = reader["student_status"].ToString();
                            addSD.DateInsert = reader["date_insert"].ToString();

                            listData.Add(addSD);
                        }
                        reader.Close();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error connecting Database: " + ex);
                }
                finally
                {
                    connect.Close();
                }
            }
            return listData;

        }

        public List<AddStudentsData> dashboardStudentData()
        {
            List<AddStudentsData> listData = new List<AddStudentsData>();

            if (connect.State != ConnectionState.Open)
            {
                try
                {
                    connect.Open();
                    DateTime today = DateTime.Today;
                    string sql = "SELECT * FROM students WHERE date_insert = @dateInsert " +
                        "AND date_delete IS NULL " +
                        "AND student_status = @studentStatus";

                    using (SqlCommand cmd = new SqlCommand(sql, connect))
                    {
                        cmd.Parameters.AddWithValue("@dateInsert", today.ToString());
                        cmd.Parameters.AddWithValue("@studentStatus", "Enrolled");

                        SqlDataReader reader = cmd.ExecuteReader();

                        while (reader.Read())
                        {
                            AddStudentsData addSD = new AddStudentsData();
                            addSD.ID = (int)reader["id"];
                            addSD.StudentID = reader["student_id"].ToString();
                            addSD.StudentName = reader["student_name"].ToString();
                            addSD.StudentGender = reader["student_gender"].ToString();
                            addSD.StudentAddress = reader["student_adress"].ToString();
                            addSD.StudentGrade = reader["student_grade"].ToString();
                            addSD.StudentSection = reader["student_section"].ToString();
                            addSD.StudentImage = reader["student_image"].ToString();
                            addSD.Status = reader["student_status"].ToString();
                            addSD.DateInsert = reader["date_insert"].ToString();

                            listData.Add(addSD);
                        }
                        reader.Close();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex);
                }
                finally
                {
                    connect.Close();
                }
            }
            return listData;
        }
    }
}
