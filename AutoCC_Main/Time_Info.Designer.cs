namespace AutoCC_Main
{
    partial class Time_Info
    {
        /// <summary> 
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary> 
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.numericUpDownEndMinute = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownEndHour = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownStartMinute = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownStartHour = new System.Windows.Forms.NumericUpDown();
            this.checkBoxActivate = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownEndMinute)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownEndHour)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStartMinute)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStartHour)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("굴림", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label2.Location = new System.Drawing.Point(3, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(386, 16);
            this.label2.TabIndex = 13;
            this.label2.Text = "------------------------------------------";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("굴림", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label1.Location = new System.Drawing.Point(37, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(25, 16);
            this.label1.TabIndex = 12;
            this.label1.Text = "일";
            // 
            // numericUpDownEndMinute
            // 
            this.numericUpDownEndMinute.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.numericUpDownEndMinute.Location = new System.Drawing.Point(335, 2);
            this.numericUpDownEndMinute.Maximum = new decimal(new int[] {
            59,
            0,
            0,
            0});
            this.numericUpDownEndMinute.Name = "numericUpDownEndMinute";
            this.numericUpDownEndMinute.Size = new System.Drawing.Size(50, 23);
            this.numericUpDownEndMinute.TabIndex = 11;
            // 
            // numericUpDownEndHour
            // 
            this.numericUpDownEndHour.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.numericUpDownEndHour.Location = new System.Drawing.Point(279, 2);
            this.numericUpDownEndHour.Maximum = new decimal(new int[] {
            23,
            0,
            0,
            0});
            this.numericUpDownEndHour.Name = "numericUpDownEndHour";
            this.numericUpDownEndHour.Size = new System.Drawing.Size(50, 23);
            this.numericUpDownEndHour.TabIndex = 10;
            // 
            // numericUpDownStartMinute
            // 
            this.numericUpDownStartMinute.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.numericUpDownStartMinute.Location = new System.Drawing.Point(198, 2);
            this.numericUpDownStartMinute.Maximum = new decimal(new int[] {
            59,
            0,
            0,
            0});
            this.numericUpDownStartMinute.Name = "numericUpDownStartMinute";
            this.numericUpDownStartMinute.Size = new System.Drawing.Size(50, 23);
            this.numericUpDownStartMinute.TabIndex = 9;
            // 
            // numericUpDownStartHour
            // 
            this.numericUpDownStartHour.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.numericUpDownStartHour.Location = new System.Drawing.Point(142, 2);
            this.numericUpDownStartHour.Maximum = new decimal(new int[] {
            23,
            0,
            0,
            0});
            this.numericUpDownStartHour.Name = "numericUpDownStartHour";
            this.numericUpDownStartHour.Size = new System.Drawing.Size(50, 23);
            this.numericUpDownStartHour.TabIndex = 8;
            // 
            // checkBoxActivate
            // 
            this.checkBoxActivate.AutoSize = true;
            this.checkBoxActivate.Location = new System.Drawing.Point(5, 6);
            this.checkBoxActivate.Name = "checkBoxActivate";
            this.checkBoxActivate.Size = new System.Drawing.Size(15, 14);
            this.checkBoxActivate.TabIndex = 7;
            this.checkBoxActivate.UseVisualStyleBackColor = true;
            this.checkBoxActivate.CheckedChanged += new System.EventHandler(this.checkBoxActivate_CheckedChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("굴림", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label3.Location = new System.Drawing.Point(253, 6);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(21, 16);
            this.label3.TabIndex = 14;
            this.label3.Text = "~";
            // 
            // Time_Info
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.numericUpDownEndMinute);
            this.Controls.Add(this.numericUpDownEndHour);
            this.Controls.Add(this.numericUpDownStartMinute);
            this.Controls.Add(this.numericUpDownStartHour);
            this.Controls.Add(this.checkBoxActivate);
            this.Controls.Add(this.label2);
            this.Name = "Time_Info";
            this.Size = new System.Drawing.Size(391, 35);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownEndMinute)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownEndHour)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStartMinute)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStartHour)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numericUpDownEndMinute;
        private System.Windows.Forms.NumericUpDown numericUpDownEndHour;
        private System.Windows.Forms.NumericUpDown numericUpDownStartMinute;
        private System.Windows.Forms.NumericUpDown numericUpDownStartHour;
        private System.Windows.Forms.CheckBox checkBoxActivate;
        private System.Windows.Forms.Label label3;
    }
}
