namespace SBL.DataLoader
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.fileDialog = new System.Windows.Forms.OpenFileDialog();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.AgentUUID = new System.Windows.Forms.TextBox();
            this.AccessToken = new System.Windows.Forms.RichTextBox();
            this.button9 = new System.Windows.Forms.Button();
            this.cbIsNeedLC = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(16, 15);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(157, 28);
            this.button1.TabIndex = 0;
            this.button1.Text = "Загрузка АС";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // fileDialog
            // 
            this.fileDialog.FileName = "openFileDialog1";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(16, 50);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(157, 28);
            this.button2.TabIndex = 1;
            this.button2.Text = "Метро";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(17, 87);
            this.button3.Margin = new System.Windows.Forms.Padding(4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(156, 28);
            this.button3.TabIndex = 2;
            this.button3.Text = "Регионы";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(16, 171);
            this.button4.Margin = new System.Windows.Forms.Padding(4);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(157, 28);
            this.button4.TabIndex = 3;
            this.button4.Text = "Метро/Регионы";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(17, 208);
            this.button5.Margin = new System.Windows.Forms.Padding(4);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(156, 28);
            this.button5.TabIndex = 4;
            this.button5.Text = "Обработка адреса";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(17, 245);
            this.button6.Margin = new System.Windows.Forms.Padding(4);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(156, 28);
            this.button6.TabIndex = 5;
            this.button6.Text = "Аптеч. сеть";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(17, 282);
            this.button7.Margin = new System.Windows.Forms.Padding(4);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(156, 28);
            this.button7.TabIndex = 6;
            this.button7.Text = "Категория";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(219, 281);
            this.button8.Margin = new System.Windows.Forms.Padding(4);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(228, 28);
            this.button8.TabIndex = 7;
            this.button8.Text = "Выгрузить Аптеки";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // AgentUUID
            // 
            this.AgentUUID.Location = new System.Drawing.Point(181, 171);
            this.AgentUUID.Margin = new System.Windows.Forms.Padding(4);
            this.AgentUUID.Name = "AgentUUID";
            this.AgentUUID.Size = new System.Drawing.Size(500, 22);
            this.AgentUUID.TabIndex = 8;
            // 
            // AccessToken
            // 
            this.AccessToken.Location = new System.Drawing.Point(183, 208);
            this.AccessToken.Margin = new System.Windows.Forms.Padding(4);
            this.AccessToken.Name = "AccessToken";
            this.AccessToken.Size = new System.Drawing.Size(499, 64);
            this.AccessToken.TabIndex = 9;
            this.AccessToken.Text = "";
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(455, 282);
            this.button9.Margin = new System.Windows.Forms.Padding(4);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(228, 28);
            this.button9.TabIndex = 10;
            this.button9.Text = "Выгрузить Сотрудников";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // cbIsNeedLC
            // 
            this.cbIsNeedLC.AutoSize = true;
            this.cbIsNeedLC.Location = new System.Drawing.Point(495, 143);
            this.cbIsNeedLC.Name = "cbIsNeedLC";
            this.cbIsNeedLC.Size = new System.Drawing.Size(192, 21);
            this.cbIsNeedLC.TabIndex = 11;
            this.cbIsNeedLC.Text = "Нужно ли создавать LC?";
            this.cbIsNeedLC.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(699, 327);
            this.Controls.Add(this.cbIsNeedLC);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.AccessToken);
            this.Controls.Add(this.AgentUUID);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "MainForm";
            this.Text = "MainForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog fileDialog;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.TextBox AgentUUID;
        private System.Windows.Forms.RichTextBox AccessToken;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.CheckBox cbIsNeedLC;
    }
}

