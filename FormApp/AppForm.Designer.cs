namespace FormApp
{
    partial class AppForm
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label_state = new System.Windows.Forms.Label();
            this.button_selectDirection = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label_state
            // 
            this.label_state.AutoSize = true;
            this.label_state.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label_state.Location = new System.Drawing.Point(12, 16);
            this.label_state.Name = "label_state";
            this.label_state.Size = new System.Drawing.Size(41, 12);
            this.label_state.TabIndex = 0;
            this.label_state.Text = "无行为";
            this.label_state.Click += new System.EventHandler(this.label_state_Click);
            // 
            // button_selectDirection
            // 
            this.button_selectDirection.Location = new System.Drawing.Point(168, 52);
            this.button_selectDirection.Name = "button_selectDirection";
            this.button_selectDirection.Size = new System.Drawing.Size(92, 23);
            this.button_selectDirection.TabIndex = 2;
            this.button_selectDirection.Text = "选择一个路径";
            this.button_selectDirection.UseVisualStyleBackColor = true;
            this.button_selectDirection.Click += new System.EventHandler(this.button_selectDirection_Click);
            // 
            // AppForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 87);
            this.Controls.Add(this.button_selectDirection);
            this.Controls.Add(this.label_state);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AppForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label_state;
        private System.Windows.Forms.Button button_selectDirection;
    }
}

