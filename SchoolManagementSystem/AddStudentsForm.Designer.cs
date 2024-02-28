namespace SchoolManagementSystem
{
    partial class AddStudentsForm
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.label1 = new System.Windows.Forms.Label();
            this.student_status = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.student_deleteBtn = new System.Windows.Forms.Button();
            this.student_clearBtn = new System.Windows.Forms.Button();
            this.student_updateBtn = new System.Windows.Forms.Button();
            this.student_addBtn = new System.Windows.Forms.Button();
            this.student_browseBtn = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.student_image = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.student_gridData = new System.Windows.Forms.DataGridView();
            this.student_section = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.student_grade = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.student_address = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.student_gender = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.student_name = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.student_id = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.student_exportBtn = new System.Windows.Forms.Button();
            this.student_importBtn = new System.Windows.Forms.Button();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.student_image)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.student_gridData)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(21, 18);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(141, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "Student\'s Data";
            // 
            // student_status
            // 
            this.student_status.FormattingEnabled = true;
            this.student_status.Items.AddRange(new object[] {
            "Enrolled",
            "Pending",
            "Graduated"});
            this.student_status.Location = new System.Drawing.Point(705, 74);
            this.student_status.Margin = new System.Windows.Forms.Padding(4);
            this.student_status.Name = "student_status";
            this.student_status.Size = new System.Drawing.Size(191, 24);
            this.student_status.TabIndex = 19;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(644, 78);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(47, 16);
            this.label8.TabIndex = 18;
            this.label8.Text = "Status:";
            // 
            // student_deleteBtn
            // 
            this.student_deleteBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.student_deleteBtn.FlatAppearance.BorderSize = 0;
            this.student_deleteBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_deleteBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_deleteBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.student_deleteBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.student_deleteBtn.ForeColor = System.Drawing.Color.White;
            this.student_deleteBtn.Location = new System.Drawing.Point(760, 214);
            this.student_deleteBtn.Margin = new System.Windows.Forms.Padding(4);
            this.student_deleteBtn.Name = "student_deleteBtn";
            this.student_deleteBtn.Size = new System.Drawing.Size(136, 43);
            this.student_deleteBtn.TabIndex = 17;
            this.student_deleteBtn.Text = "Delete";
            this.student_deleteBtn.UseVisualStyleBackColor = false;
            this.student_deleteBtn.Click += new System.EventHandler(this.student_deleteBtn_Click);
            // 
            // student_clearBtn
            // 
            this.student_clearBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.student_clearBtn.FlatAppearance.BorderSize = 0;
            this.student_clearBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_clearBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_clearBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.student_clearBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.student_clearBtn.ForeColor = System.Drawing.Color.White;
            this.student_clearBtn.Location = new System.Drawing.Point(589, 214);
            this.student_clearBtn.Margin = new System.Windows.Forms.Padding(4);
            this.student_clearBtn.Name = "student_clearBtn";
            this.student_clearBtn.Size = new System.Drawing.Size(136, 43);
            this.student_clearBtn.TabIndex = 16;
            this.student_clearBtn.Text = "Clear";
            this.student_clearBtn.UseVisualStyleBackColor = false;
            this.student_clearBtn.Click += new System.EventHandler(this.student_clearBtn_Click);
            // 
            // student_updateBtn
            // 
            this.student_updateBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.student_updateBtn.FlatAppearance.BorderSize = 0;
            this.student_updateBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_updateBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_updateBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.student_updateBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.student_updateBtn.ForeColor = System.Drawing.Color.White;
            this.student_updateBtn.Location = new System.Drawing.Point(416, 214);
            this.student_updateBtn.Margin = new System.Windows.Forms.Padding(4);
            this.student_updateBtn.Name = "student_updateBtn";
            this.student_updateBtn.Size = new System.Drawing.Size(136, 43);
            this.student_updateBtn.TabIndex = 15;
            this.student_updateBtn.Text = "Update";
            this.student_updateBtn.UseVisualStyleBackColor = false;
            this.student_updateBtn.Click += new System.EventHandler(this.student_updateBtn_Click);
            // 
            // student_addBtn
            // 
            this.student_addBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.student_addBtn.FlatAppearance.BorderSize = 0;
            this.student_addBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_addBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_addBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.student_addBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.student_addBtn.ForeColor = System.Drawing.Color.White;
            this.student_addBtn.Location = new System.Drawing.Point(248, 214);
            this.student_addBtn.Margin = new System.Windows.Forms.Padding(4);
            this.student_addBtn.Name = "student_addBtn";
            this.student_addBtn.Size = new System.Drawing.Size(136, 43);
            this.student_addBtn.TabIndex = 14;
            this.student_addBtn.Text = "Add";
            this.student_addBtn.UseVisualStyleBackColor = false;
            this.student_addBtn.Click += new System.EventHandler(this.student_addBtn_Click);
            // 
            // student_browseBtn
            // 
            this.student_browseBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.student_browseBtn.FlatAppearance.BorderSize = 0;
            this.student_browseBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_browseBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_browseBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.student_browseBtn.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.student_browseBtn.ForeColor = System.Drawing.Color.White;
            this.student_browseBtn.Location = new System.Drawing.Point(1000, 132);
            this.student_browseBtn.Margin = new System.Windows.Forms.Padding(4);
            this.student_browseBtn.Name = "student_browseBtn";
            this.student_browseBtn.Size = new System.Drawing.Size(107, 28);
            this.student_browseBtn.TabIndex = 13;
            this.student_browseBtn.Text = "Browse";
            this.student_browseBtn.UseVisualStyleBackColor = false;
            this.student_browseBtn.Click += new System.EventHandler(this.student_browseBtn_Click);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.Silver;
            this.panel3.Controls.Add(this.student_image);
            this.panel3.Location = new System.Drawing.Point(1000, 30);
            this.panel3.Margin = new System.Windows.Forms.Padding(4);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(107, 106);
            this.panel3.TabIndex = 12;
            // 
            // student_image
            // 
            this.student_image.Location = new System.Drawing.Point(0, 0);
            this.student_image.Margin = new System.Windows.Forms.Padding(4);
            this.student_image.Name = "student_image";
            this.student_image.Size = new System.Drawing.Size(107, 106);
            this.student_image.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.student_image.TabIndex = 1;
            this.student_image.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.student_gridData);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(16, 25);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1134, 356);
            this.panel1.TabIndex = 4;
            // 
            // student_gridData
            // 
            this.student_gridData.AllowUserToAddRows = false;
            this.student_gridData.AllowUserToDeleteRows = false;
            this.student_gridData.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.student_gridData.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.student_gridData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.student_gridData.EnableHeadersVisualStyles = false;
            this.student_gridData.Location = new System.Drawing.Point(27, 58);
            this.student_gridData.Margin = new System.Windows.Forms.Padding(4);
            this.student_gridData.Name = "student_gridData";
            this.student_gridData.ReadOnly = true;
            this.student_gridData.RowHeadersVisible = false;
            this.student_gridData.RowHeadersWidth = 51;
            this.student_gridData.Size = new System.Drawing.Size(1080, 276);
            this.student_gridData.TabIndex = 1;
            this.student_gridData.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.student_gridData_CellClick);
            // 
            // student_section
            // 
            this.student_section.FormattingEnabled = true;
            this.student_section.Items.AddRange(new object[] {
            "A",
            "B",
            "C",
            "D",
            "E"});
            this.student_section.Location = new System.Drawing.Point(705, 26);
            this.student_section.Margin = new System.Windows.Forms.Padding(4);
            this.student_section.Name = "student_section";
            this.student_section.Size = new System.Drawing.Size(191, 24);
            this.student_section.TabIndex = 11;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(636, 30);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(55, 16);
            this.label7.TabIndex = 10;
            this.label7.Text = "Section:";
            // 
            // student_grade
            // 
            this.student_grade.FormattingEnabled = true;
            this.student_grade.Items.AddRange(new object[] {
            "Grade A",
            "Grade A+",
            "Grade B",
            "Grade B+",
            "Grade C",
            "Grade C+"});
            this.student_grade.Location = new System.Drawing.Point(416, 122);
            this.student_grade.Margin = new System.Windows.Forms.Padding(4);
            this.student_grade.Name = "student_grade";
            this.student_grade.Size = new System.Drawing.Size(191, 24);
            this.student_grade.TabIndex = 9;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(356, 126);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 16);
            this.label6.TabIndex = 8;
            this.label6.Text = "Grade:";
            // 
            // student_address
            // 
            this.student_address.Location = new System.Drawing.Point(416, 26);
            this.student_address.Margin = new System.Windows.Forms.Padding(4);
            this.student_address.Multiline = true;
            this.student_address.Name = "student_address";
            this.student_address.Size = new System.Drawing.Size(191, 80);
            this.student_address.TabIndex = 7;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(344, 33);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(61, 16);
            this.label5.TabIndex = 6;
            this.label5.Text = "Address:";
            // 
            // student_gender
            // 
            this.student_gender.FormattingEnabled = true;
            this.student_gender.Items.AddRange(new object[] {
            "Male",
            "Female",
            "Others"});
            this.student_gender.Location = new System.Drawing.Point(108, 128);
            this.student_gender.Margin = new System.Windows.Forms.Padding(4);
            this.student_gender.Name = "student_gender";
            this.student_gender.Size = new System.Drawing.Size(191, 24);
            this.student_gender.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(40, 132);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 16);
            this.label4.TabIndex = 4;
            this.label4.Text = "Gender:";
            // 
            // student_name
            // 
            this.student_name.Location = new System.Drawing.Point(107, 76);
            this.student_name.Margin = new System.Windows.Forms.Padding(4);
            this.student_name.Multiline = true;
            this.student_name.Name = "student_name";
            this.student_name.Size = new System.Drawing.Size(191, 30);
            this.student_name.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 84);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 16);
            this.label3.TabIndex = 2;
            this.label3.Text = "Full Name:";
            // 
            // student_id
            // 
            this.student_id.Location = new System.Drawing.Point(107, 26);
            this.student_id.Margin = new System.Windows.Forms.Padding(4);
            this.student_id.Multiline = true;
            this.student_id.Name = "student_id";
            this.student_id.Size = new System.Drawing.Size(135, 30);
            this.student_id.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 33);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 16);
            this.label2.TabIndex = 0;
            this.label2.Text = "Student ID:";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.White;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.student_exportBtn);
            this.panel2.Controls.Add(this.student_importBtn);
            this.panel2.Controls.Add(this.student_status);
            this.panel2.Controls.Add(this.label8);
            this.panel2.Controls.Add(this.student_deleteBtn);
            this.panel2.Controls.Add(this.student_clearBtn);
            this.panel2.Controls.Add(this.student_updateBtn);
            this.panel2.Controls.Add(this.student_addBtn);
            this.panel2.Controls.Add(this.student_browseBtn);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Controls.Add(this.student_section);
            this.panel2.Controls.Add(this.label7);
            this.panel2.Controls.Add(this.student_grade);
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.student_address);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.student_gender);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.student_name);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.student_id);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Location = new System.Drawing.Point(16, 403);
            this.panel2.Margin = new System.Windows.Forms.Padding(4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1134, 281);
            this.panel2.TabIndex = 5;
            // 
            // student_exportBtn
            // 
            this.student_exportBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(116)))), ((int)(((byte)(59)))));
            this.student_exportBtn.FlatAppearance.BorderSize = 0;
            this.student_exportBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_exportBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_exportBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.student_exportBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.student_exportBtn.ForeColor = System.Drawing.Color.White;
            this.student_exportBtn.Location = new System.Drawing.Point(1035, 214);
            this.student_exportBtn.Margin = new System.Windows.Forms.Padding(4);
            this.student_exportBtn.Name = "student_exportBtn";
            this.student_exportBtn.Size = new System.Drawing.Size(72, 43);
            this.student_exportBtn.TabIndex = 21;
            this.student_exportBtn.Text = "Export";
            this.student_exportBtn.UseVisualStyleBackColor = false;
            this.student_exportBtn.Click += new System.EventHandler(this.student_exportBtn_Click);
            // 
            // student_importBtn
            // 
            this.student_importBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(116)))), ((int)(((byte)(59)))));
            this.student_importBtn.FlatAppearance.BorderSize = 0;
            this.student_importBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_importBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.student_importBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.student_importBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.student_importBtn.ForeColor = System.Drawing.Color.White;
            this.student_importBtn.Location = new System.Drawing.Point(929, 214);
            this.student_importBtn.Margin = new System.Windows.Forms.Padding(4);
            this.student_importBtn.Name = "student_importBtn";
            this.student_importBtn.Size = new System.Drawing.Size(76, 43);
            this.student_importBtn.TabIndex = 20;
            this.student_importBtn.Text = "Import";
            this.student_importBtn.UseVisualStyleBackColor = false;
            this.student_importBtn.Click += new System.EventHandler(this.student_importBtn_Click);
            // 
            // AddStudentsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Name = "AddStudentsForm";
            this.Size = new System.Drawing.Size(1167, 708);
            this.panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.student_image)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.student_gridData)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox student_status;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button student_deleteBtn;
        private System.Windows.Forms.Button student_clearBtn;
        private System.Windows.Forms.Button student_updateBtn;
        private System.Windows.Forms.Button student_addBtn;
        private System.Windows.Forms.Button student_browseBtn;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.PictureBox student_image;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView student_gridData;
        private System.Windows.Forms.ComboBox student_section;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox student_grade;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox student_address;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox student_gender;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox student_name;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox student_id;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button student_exportBtn;
        private System.Windows.Forms.Button student_importBtn;
    }
}
