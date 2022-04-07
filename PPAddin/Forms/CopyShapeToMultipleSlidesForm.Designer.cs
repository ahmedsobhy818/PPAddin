
namespace PPAddin.Forms
{
    partial class CopyShapeToMultipleSlidesForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.ShapeIdentifierTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.AllSlidesListBox = new System.Windows.Forms.ListBox();
            this.OptionExistingShapes1 = new System.Windows.Forms.RadioButton();
            this.OptionExistingShapes2 = new System.Windows.Forms.RadioButton();
            this.CopyShapesToSelectedSlidesButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Shape identifier";
            // 
            // ShapeIdentifierTextBox
            // 
            this.ShapeIdentifierTextBox.Location = new System.Drawing.Point(156, 22);
            this.ShapeIdentifierTextBox.Name = "ShapeIdentifierTextBox";
            this.ShapeIdentifierTextBox.Size = new System.Drawing.Size(175, 24);
            this.ShapeIdentifierTextBox.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(362, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(176, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "Pick a unique identifie/name";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(31, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(243, 17);
            this.label3.TabIndex = 3;
            this.label3.Text = "Copy this shape to the following slides:";
            // 
            // AllSlidesListBox
            // 
            this.AllSlidesListBox.FormattingEnabled = true;
            this.AllSlidesListBox.ItemHeight = 16;
            this.AllSlidesListBox.Location = new System.Drawing.Point(34, 130);
            this.AllSlidesListBox.Name = "AllSlidesListBox";
            this.AllSlidesListBox.Size = new System.Drawing.Size(642, 260);
            this.AllSlidesListBox.TabIndex = 4;
            // 
            // OptionExistingShapes1
            // 
            this.OptionExistingShapes1.AutoSize = true;
            this.OptionExistingShapes1.Checked = true;
            this.OptionExistingShapes1.Location = new System.Drawing.Point(34, 417);
            this.OptionExistingShapes1.Name = "OptionExistingShapes1";
            this.OptionExistingShapes1.Size = new System.Drawing.Size(184, 21);
            this.OptionExistingShapes1.TabIndex = 5;
            this.OptionExistingShapes1.TabStop = true;
            this.OptionExistingShapes1.Text = "Overwrite existing shapes";
            this.OptionExistingShapes1.UseVisualStyleBackColor = true;
            // 
            // OptionExistingShapes2
            // 
            this.OptionExistingShapes2.AutoSize = true;
            this.OptionExistingShapes2.Location = new System.Drawing.Point(230, 417);
            this.OptionExistingShapes2.Name = "OptionExistingShapes2";
            this.OptionExistingShapes2.Size = new System.Drawing.Size(214, 21);
            this.OptionExistingShapes2.TabIndex = 6;
            this.OptionExistingShapes2.TabStop = true;
            this.OptionExistingShapes2.Text = "Skip slides with existing shapes";
            this.OptionExistingShapes2.UseVisualStyleBackColor = true;
            // 
            // CopyShapesToSelectedSlidesButton
            // 
            this.CopyShapesToSelectedSlidesButton.Location = new System.Drawing.Point(460, 405);
            this.CopyShapesToSelectedSlidesButton.Name = "CopyShapesToSelectedSlidesButton";
            this.CopyShapesToSelectedSlidesButton.Size = new System.Drawing.Size(216, 42);
            this.CopyShapesToSelectedSlidesButton.TabIndex = 7;
            this.CopyShapesToSelectedSlidesButton.Text = "Copy shape to selected slides";
            this.CopyShapesToSelectedSlidesButton.UseVisualStyleBackColor = true;
            this.CopyShapesToSelectedSlidesButton.Click += new System.EventHandler(this.CopyShapesToSelectedSlidesButton_Click);
            // 
            // CopyShapeToMultipleSlidesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.CopyShapesToSelectedSlidesButton);
            this.Controls.Add(this.OptionExistingShapes2);
            this.Controls.Add(this.OptionExistingShapes1);
            this.Controls.Add(this.AllSlidesListBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ShapeIdentifierTextBox);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CopyShapeToMultipleSlidesForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Copy shape to multiple slides";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        public  System.Windows.Forms.TextBox ShapeIdentifierTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        public  System.Windows.Forms.ListBox AllSlidesListBox;
        private System.Windows.Forms.RadioButton OptionExistingShapes1;
        private System.Windows.Forms.RadioButton OptionExistingShapes2;
        private System.Windows.Forms.Button CopyShapesToSelectedSlidesButton;
    }
}