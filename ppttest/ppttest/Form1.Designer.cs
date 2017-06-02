namespace ppttest
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_create = new System.Windows.Forms.Button();
            this.btn_merge = new System.Windows.Forms.Button();
            this.btn_split = new System.Windows.Forms.Button();
            this.btn_findshape = new System.Windows.Forms.Button();
            this.btn_singlegenarate = new System.Windows.Forms.Button();
            this.btn_findhidden = new System.Windows.Forms.Button();
            this.btn_mutigenerate = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_create
            // 
            this.btn_create.Location = new System.Drawing.Point(0, 0);
            this.btn_create.Name = "btn_create";
            this.btn_create.Size = new System.Drawing.Size(75, 23);
            this.btn_create.TabIndex = 0;
            this.btn_create.Text = "创建ppt";
            this.btn_create.UseVisualStyleBackColor = true;
            this.btn_create.Click += new System.EventHandler(this.btn_create_Click);
            // 
            // btn_merge
            // 
            this.btn_merge.Location = new System.Drawing.Point(165, 181);
            this.btn_merge.Name = "btn_merge";
            this.btn_merge.Size = new System.Drawing.Size(75, 23);
            this.btn_merge.TabIndex = 1;
            this.btn_merge.Text = "合并ppt";
            this.btn_merge.UseVisualStyleBackColor = true;
            this.btn_merge.Click += new System.EventHandler(this.btn_merge_Click);
            // 
            // btn_split
            // 
            this.btn_split.Location = new System.Drawing.Point(246, 181);
            this.btn_split.Name = "btn_split";
            this.btn_split.Size = new System.Drawing.Size(75, 23);
            this.btn_split.TabIndex = 2;
            this.btn_split.Text = "拆分ppt";
            this.btn_split.UseVisualStyleBackColor = true;
            this.btn_split.Click += new System.EventHandler(this.btn_split_Click);
            // 
            // btn_findshape
            // 
            this.btn_findshape.Location = new System.Drawing.Point(103, 0);
            this.btn_findshape.Name = "btn_findshape";
            this.btn_findshape.Size = new System.Drawing.Size(75, 23);
            this.btn_findshape.TabIndex = 3;
            this.btn_findshape.Text = "查询隐藏的shape";
            this.btn_findshape.UseVisualStyleBackColor = true;
            this.btn_findshape.Click += new System.EventHandler(this.btn_findshape_Click);
            // 
            // btn_singlegenarate
            // 
            this.btn_singlegenarate.Location = new System.Drawing.Point(0, 128);
            this.btn_singlegenarate.Name = "btn_singlegenarate";
            this.btn_singlegenarate.Size = new System.Drawing.Size(159, 23);
            this.btn_singlegenarate.TabIndex = 4;
            this.btn_singlegenarate.Text = "根据模板生成单个ppt";
            this.btn_singlegenarate.UseVisualStyleBackColor = true;
            this.btn_singlegenarate.Click += new System.EventHandler(this.btn_singlegenarate_Click);
            // 
            // btn_findhidden
            // 
            this.btn_findhidden.Location = new System.Drawing.Point(165, 128);
            this.btn_findhidden.Name = "btn_findhidden";
            this.btn_findhidden.Size = new System.Drawing.Size(115, 23);
            this.btn_findhidden.TabIndex = 5;
            this.btn_findhidden.Text = "查找隐藏元素";
            this.btn_findhidden.UseVisualStyleBackColor = true;
            this.btn_findhidden.Click += new System.EventHandler(this.btn_findhidden_Click);
            // 
            // btn_mutigenerate
            // 
            this.btn_mutigenerate.Location = new System.Drawing.Point(0, 181);
            this.btn_mutigenerate.Name = "btn_mutigenerate";
            this.btn_mutigenerate.Size = new System.Drawing.Size(159, 23);
            this.btn_mutigenerate.TabIndex = 6;
            this.btn_mutigenerate.Text = "根据模板生成多个ppt";
            this.btn_mutigenerate.UseVisualStyleBackColor = true;
            this.btn_mutigenerate.Click += new System.EventHandler(this.btn_mutigenerate_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(582, 262);
            this.Controls.Add(this.btn_mutigenerate);
            this.Controls.Add(this.btn_findhidden);
            this.Controls.Add(this.btn_singlegenarate);
            this.Controls.Add(this.btn_findshape);
            this.Controls.Add(this.btn_split);
            this.Controls.Add(this.btn_merge);
            this.Controls.Add(this.btn_create);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_create;
        private System.Windows.Forms.Button btn_merge;
        private System.Windows.Forms.Button btn_split;
        private System.Windows.Forms.Button btn_findshape;
        private System.Windows.Forms.Button btn_singlegenarate;
        private System.Windows.Forms.Button btn_findhidden;
        private System.Windows.Forms.Button btn_mutigenerate;
    }
}

