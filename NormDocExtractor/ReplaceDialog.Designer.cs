namespace NormDocExtractor
{
  partial class ReplaceDialog
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
      this.lblSource = new System.Windows.Forms.Label();
      this.txtSource = new System.Windows.Forms.TextBox();
      this.lblOld = new System.Windows.Forms.Label();
      this.txtOld = new System.Windows.Forms.TextBox();
      this.lblNew = new System.Windows.Forms.Label();
      this.txtNew = new System.Windows.Forms.TextBox();
      this.lblCount = new System.Windows.Forms.Label();
      this.btnCount = new System.Windows.Forms.Button();
      this.btnReplace = new System.Windows.Forms.Button();
      this.btnCancel = new System.Windows.Forms.Button();
      this.SuspendLayout();
      // 
      // lblSource
      // 
      this.lblSource.AutoSize = true;
      this.lblSource.Location = new System.Drawing.Point(12, 9);
      this.lblSource.Name = "lblSource";
      this.lblSource.Size = new System.Drawing.Size(112, 13);
      this.lblSource.TabIndex = 0;
      this.lblSource.Text = "Исходный документ:";
      // 
      // txtSource
      // 
      this.txtSource.Location = new System.Drawing.Point(12, 25);
      this.txtSource.Name = "txtSource";
      this.txtSource.ReadOnly = true;
      this.txtSource.Size = new System.Drawing.Size(509, 20);
      this.txtSource.TabIndex = 1;
      // 
      // lblOld
      // 
      this.lblOld.AutoSize = true;
      this.lblOld.Location = new System.Drawing.Point(13, 64);
      this.lblOld.Name = "lblOld";
      this.lblOld.Size = new System.Drawing.Size(107, 13);
      this.lblOld.TabIndex = 2;
      this.lblOld.Text = "Найти в документе:";
      // 
      // txtOld
      // 
      this.txtOld.Location = new System.Drawing.Point(12, 81);
      this.txtOld.MaxLength = 200;
      this.txtOld.Name = "txtOld";
      this.txtOld.Size = new System.Drawing.Size(509, 20);
      this.txtOld.TabIndex = 3;
      // 
      // lblNew
      // 
      this.lblNew.AutoSize = true;
      this.lblNew.Location = new System.Drawing.Point(13, 121);
      this.lblNew.Name = "lblNew";
      this.lblNew.Size = new System.Drawing.Size(75, 13);
      this.lblNew.TabIndex = 4;
      this.lblNew.Text = "Заменить на:";
      // 
      // txtNew
      // 
      this.txtNew.Location = new System.Drawing.Point(12, 138);
      this.txtNew.MaxLength = 200;
      this.txtNew.Name = "txtNew";
      this.txtNew.Size = new System.Drawing.Size(509, 20);
      this.txtNew.TabIndex = 5;
      // 
      // lblCount
      // 
      this.lblCount.AutoSize = true;
      this.lblCount.Location = new System.Drawing.Point(16, 188);
      this.lblCount.Name = "lblCount";
      this.lblCount.Size = new System.Drawing.Size(63, 13);
      this.lblCount.TabIndex = 6;
      this.lblCount.Text = "Найдено: —";
      // 
      // btnCount
      // 
      this.btnCount.Location = new System.Drawing.Point(110, 183);
      this.btnCount.Name = "btnCount";
      this.btnCount.Size = new System.Drawing.Size(75, 23);
      this.btnCount.TabIndex = 7;
      this.btnCount.Text = "Посчитать";
      this.btnCount.UseVisualStyleBackColor = true;
      // 
      // btnReplace
      // 
      this.btnReplace.Location = new System.Drawing.Point(333, 183);
      this.btnReplace.Name = "btnReplace";
      this.btnReplace.Size = new System.Drawing.Size(104, 23);
      this.btnReplace.TabIndex = 8;
      this.btnReplace.Text = "Заменить всё";
      this.btnReplace.UseVisualStyleBackColor = true;
      // 
      // btnCancel
      // 
      this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
      this.btnCancel.Location = new System.Drawing.Point(443, 183);
      this.btnCancel.Name = "btnCancel";
      this.btnCancel.Size = new System.Drawing.Size(75, 23);
      this.btnCancel.TabIndex = 9;
      this.btnCancel.Text = "Отмена";
      this.btnCancel.UseVisualStyleBackColor = true;
      // 
      // ReplaceDialog
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.CancelButton = this.btnCancel;
      this.ClientSize = new System.Drawing.Size(531, 219);
      this.Controls.Add(this.btnCancel);
      this.Controls.Add(this.btnReplace);
      this.Controls.Add(this.btnCount);
      this.Controls.Add(this.lblCount);
      this.Controls.Add(this.txtNew);
      this.Controls.Add(this.lblNew);
      this.Controls.Add(this.txtOld);
      this.Controls.Add(this.lblOld);
      this.Controls.Add(this.txtSource);
      this.Controls.Add(this.lblSource);
      this.Name = "ReplaceDialog";
      this.Text = "NormDocExtract — Замена документа в исходнике";
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.Label lblSource;
    private System.Windows.Forms.TextBox txtSource;
    private System.Windows.Forms.Label lblOld;
    private System.Windows.Forms.TextBox txtOld;
    private System.Windows.Forms.Label lblNew;
    private System.Windows.Forms.TextBox txtNew;
    private System.Windows.Forms.Label lblCount;
    private System.Windows.Forms.Button btnCount;
    private System.Windows.Forms.Button btnReplace;
    private System.Windows.Forms.Button btnCancel;
  }
}