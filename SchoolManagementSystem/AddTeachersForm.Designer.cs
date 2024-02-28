namespace SchoolManagementSystem
{
    partial class AddTeachersForm
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.teacher_gridData = new System.Windows.Forms.DataGridView();
            this.teacher_deleteBtn = new System.Windows.Forms.Button();
            this.teacher_clearBtn = new System.Windows.Forms.Button();
            this.teacher_updateBtn = new System.Windows.Forms.Button();
            this.teacher_addBtn = new System.Windows.Forms.Button();
            this.teacher_browseBtn = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.teacher_image = new System.Windows.Forms.PictureBox();
            this.teacher_status = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.teacher_exportBtn = new System.Windows.Forms.Button();
            this.teacher_importBtn = new System.Windows.Forms.Button();
            this.teacher_address = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.teacher_gender = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.teacher_name = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.teacher_id = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.teacher_gridData)).BeginInit();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.teacher_image)).BeginInit();
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
            this.label1.Size = new System.Drawing.Size(144, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "Teacher\'s Data";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.teacher_gridData);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(16, 25);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1134, 356);
            this.panel1.TabIndex = 6;
            // 
            // teacher_gridData
            // 
            this.teacher_gridData.AllowUserToAddRows = false;
            this.teacher_gridData.AllowUserToDeleteRows = false;
            this.teacher_gridData.AllowUserToResizeRows = false;
            this.teacher_gridData.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.teacher_gridData.CausesValidation = false;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.teacher_gridData.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.teacher_gridData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.teacher_gridData.EnableHeadersVisualStyles = false;
            this.teacher_gridData.Location = new System.Drawing.Point(27, 58);
            this.teacher_gridData.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_gridData.Name = "teacher_gridData";
            this.teacher_gridData.ReadOnly = true;
            this.teacher_gridData.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.teacher_gridData.RowHeadersVisible = false;
            this.teacher_gridData.RowHeadersWidth = 51;
            this.teacher_gridData.Size = new System.Drawing.Size(1080, 276);
            this.teacher_gridData.TabIndex = 1;
            this.teacher_gridData.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.teacher_gridData_CellClick);
            // 
            // teacher_deleteBtn
            // 
            this.teacher_deleteBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.teacher_deleteBtn.FlatAppearance.BorderSize = 0;
            this.teacher_deleteBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_deleteBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_deleteBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.teacher_deleteBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.teacher_deleteBtn.ForeColor = System.Drawing.Color.White;
            this.teacher_deleteBtn.Location = new System.Drawing.Point(760, 214);
            this.teacher_deleteBtn.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_deleteBtn.Name = "teacher_deleteBtn";
            this.teacher_deleteBtn.Size = new System.Drawing.Size(136, 43);
            this.teacher_deleteBtn.TabIndex = 17;
            this.teacher_deleteBtn.Text = "Delete";
            this.teacher_deleteBtn.UseVisualStyleBackColor = false;
            this.teacher_deleteBtn.Click += new System.EventHandler(this.teacher_deleteBtn_Click);
            // 
            // teacher_clearBtn
            // 
            this.teacher_clearBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.teacher_clearBtn.FlatAppearance.BorderSize = 0;
            this.teacher_clearBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_clearBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_clearBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.teacher_clearBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.teacher_clearBtn.ForeColor = System.Drawing.Color.White;
            this.teacher_clearBtn.Location = new System.Drawing.Point(587, 214);
            this.teacher_clearBtn.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_clearBtn.Name = "teacher_clearBtn";
            this.teacher_clearBtn.Size = new System.Drawing.Size(136, 43);
            this.teacher_clearBtn.TabIndex = 16;
            this.teacher_clearBtn.Text = "Clear";
            this.teacher_clearBtn.UseVisualStyleBackColor = false;
            this.teacher_clearBtn.Click += new System.EventHandler(this.teacher_clearBtn_Click);
            // 
            // teacher_updateBtn
            // 
            this.teacher_updateBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.teacher_updateBtn.FlatAppearance.BorderSize = 0;
            this.teacher_updateBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_updateBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_updateBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.teacher_updateBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.teacher_updateBtn.ForeColor = System.Drawing.Color.White;
            this.teacher_updateBtn.Location = new System.Drawing.Point(416, 214);
            this.teacher_updateBtn.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_updateBtn.Name = "teacher_updateBtn";
            this.teacher_updateBtn.Size = new System.Drawing.Size(136, 43);
            this.teacher_updateBtn.TabIndex = 15;
            this.teacher_updateBtn.Text = "Update";
            this.teacher_updateBtn.UseVisualStyleBackColor = false;
            this.teacher_updateBtn.Click += new System.EventHandler(this.teacher_updateBtn_Click);
            // 
            // teacher_addBtn
            // 
            this.teacher_addBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.teacher_addBtn.FlatAppearance.BorderSize = 0;
            this.teacher_addBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_addBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_addBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.teacher_addBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.teacher_addBtn.ForeColor = System.Drawing.Color.White;
            this.teacher_addBtn.Location = new System.Drawing.Point(248, 214);
            this.teacher_addBtn.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_addBtn.Name = "teacher_addBtn";
            this.teacher_addBtn.Size = new System.Drawing.Size(136, 43);
            this.teacher_addBtn.TabIndex = 14;
            this.teacher_addBtn.Text = "Add";
            this.teacher_addBtn.UseVisualStyleBackColor = false;
            this.teacher_addBtn.Click += new System.EventHandler(this.teacher_addBtn_Click);
            // 
            // teacher_browseBtn
            // 
            this.teacher_browseBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(87)))), ((int)(((byte)(122)))));
            this.teacher_browseBtn.FlatAppearance.BorderSize = 0;
            this.teacher_browseBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_browseBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_browseBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.teacher_browseBtn.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.teacher_browseBtn.ForeColor = System.Drawing.Color.White;
            this.teacher_browseBtn.Location = new System.Drawing.Point(1000, 132);
            this.teacher_browseBtn.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_browseBtn.Name = "teacher_browseBtn";
            this.teacher_browseBtn.Size = new System.Drawing.Size(107, 28);
            this.teacher_browseBtn.TabIndex = 13;
            this.teacher_browseBtn.Text = "Browse";
            this.teacher_browseBtn.UseVisualStyleBackColor = false;
            this.teacher_browseBtn.Click += new System.EventHandler(this.teacher_browseBtn_Click);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.Silver;
            this.panel3.Controls.Add(this.teacher_image);
            this.panel3.Location = new System.Drawing.Point(1000, 30);
            this.panel3.Margin = new System.Windows.Forms.Padding(4);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(107, 106);
            this.panel3.TabIndex = 12;
            // 
            // teacher_image
            // 
            this.teacher_image.Location = new System.Drawing.Point(0, 0);
            this.teacher_image.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_image.Name = "teacher_image";
            this.teacher_image.Size = new System.Drawing.Size(107, 106);
            this.teacher_image.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.teacher_image.TabIndex = 0;
            this.teacher_image.TabStop = false;
            // 
            // teacher_status
            // 
            this.teacher_status.FormattingEnabled = true;
            this.teacher_status.Items.AddRange(new object[] {
            "Active",
            "Inactive"});
            this.teacher_status.Location = new System.Drawing.Point(705, 26);
            this.teacher_status.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_status.Name = "teacher_status";
            this.teacher_status.Size = new System.Drawing.Size(191, 24);
            this.teacher_status.TabIndex = 11;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(644, 30);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(47, 16);
            this.label7.TabIndex = 10;
            this.label7.Text = "Status:";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.White;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.teacher_exportBtn);
            this.panel2.Controls.Add(this.teacher_importBtn);
            this.panel2.Controls.Add(this.teacher_deleteBtn);
            this.panel2.Controls.Add(this.teacher_clearBtn);
            this.panel2.Controls.Add(this.teacher_updateBtn);
            this.panel2.Controls.Add(this.teacher_addBtn);
            this.panel2.Controls.Add(this.teacher_browseBtn);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Controls.Add(this.teacher_status);
            this.panel2.Controls.Add(this.label7);
            this.panel2.Controls.Add(this.teacher_address);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.teacher_gender);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.teacher_name);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.teacher_id);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Location = new System.Drawing.Point(16, 403);
            this.panel2.Margin = new System.Windows.Forms.Padding(4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1134, 281);
            this.panel2.TabIndex = 7;
            // 
            // teacher_exportBtn
            // 
            this.teacher_exportBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(116)))), ((int)(((byte)(59)))));
            this.teacher_exportBtn.FlatAppearance.BorderSize = 0;
            this.teacher_exportBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_exportBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_exportBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.teacher_exportBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.teacher_exportBtn.ForeColor = System.Drawing.Color.White;
            this.teacher_exportBtn.Location = new System.Drawing.Point(1035, 214);
            this.teacher_exportBtn.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_exportBtn.Name = "teacher_exportBtn";
            this.teacher_exportBtn.Size = new System.Drawing.Size(72, 43);
            this.teacher_exportBtn.TabIndex = 19;
            this.teacher_exportBtn.Text = "Export";
            this.teacher_exportBtn.UseVisualStyleBackColor = false;
            this.teacher_exportBtn.Click += new System.EventHandler(this.teacher_exportBtn_Click);
            // 
            // teacher_importBtn
            // 
            this.teacher_importBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(116)))), ((int)(((byte)(59)))));
            this.teacher_importBtn.FlatAppearance.BorderSize = 0;
            this.teacher_importBtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_importBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.teacher_importBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.teacher_importBtn.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.teacher_importBtn.ForeColor = System.Drawing.Color.White;
            this.teacher_importBtn.Location = new System.Drawing.Point(929, 214);
            this.teacher_importBtn.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_importBtn.Name = "teacher_importBtn";
            this.teacher_importBtn.Size = new System.Drawing.Size(76, 43);
            this.teacher_importBtn.TabIndex = 18;
            this.teacher_importBtn.Text = "Import";
            this.teacher_importBtn.UseVisualStyleBackColor = false;
            this.teacher_importBtn.Click += new System.EventHandler(this.teacher_importBtn_Click);
            // 
            // teacher_address
            // 
            this.teacher_address.Location = new System.Drawing.Point(416, 26);
            this.teacher_address.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_address.Multiline = true;
            this.teacher_address.Name = "teacher_address";
            this.teacher_address.Size = new System.Drawing.Size(191, 80);
            this.teacher_address.TabIndex = 7;
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
            // teacher_gender
            // 
            this.teacher_gender.FormattingEnabled = true;
            this.teacher_gender.Items.AddRange(new object[] {
            "Male",
            "Female",
            "Others"});
            this.teacher_gender.Location = new System.Drawing.Point(108, 128);
            this.teacher_gender.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_gender.Name = "teacher_gender";
            this.teacher_gender.Size = new System.Drawing.Size(191, 24);
            this.teacher_gender.TabIndex = 5;
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
            // teacher_name
            // 
            this.teacher_name.Location = new System.Drawing.Point(107, 76);
            this.teacher_name.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_name.Multiline = true;
            this.teacher_name.Name = "teacher_name";
            this.teacher_name.Size = new System.Drawing.Size(191, 30);
            this.teacher_name.TabIndex = 3;
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
            // teacher_id
            // 
            this.teacher_id.Location = new System.Drawing.Point(107, 26);
            this.teacher_id.Margin = new System.Windows.Forms.Padding(4);
            this.teacher_id.Multiline = true;
            this.teacher_id.Name = "teacher_id";
            this.teacher_id.Size = new System.Drawing.Size(135, 30);
            this.teacher_id.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 33);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 16);
            this.label2.TabIndex = 0;
            this.label2.Text = "Teacher ID:";
            // 
            // AddTeachersForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Name = "AddTeachersForm";
            this.Size = new System.Drawing.Size(1167, 708);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.teacher_gridData)).EndInit();
            this.panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.teacher_image)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView teacher_gridData;
        private System.Windows.Forms.Button teacher_deleteBtn;
        private System.Windows.Forms.Button teacher_clearBtn;
        private System.Windows.Forms.Button teacher_updateBtn;
        private System.Windows.Forms.Button teacher_addBtn;
        private System.Windows.Forms.Button teacher_browseBtn;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.PictureBox teacher_image;
        private System.Windows.Forms.ComboBox teacher_status;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox teacher_address;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox teacher_gender;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox teacher_name;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox teacher_id;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button teacher_exportBtn;
        private System.Windows.Forms.Button teacher_importBtn;
    }
}
