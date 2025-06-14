namespace aposplit
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            lblFichier = new Label();
            toolTip1 = new ToolTip(components);
            pnlStatusBar = new Panel();
            lblInfo = new Label();
            SuspendLayout();
            // 
            // lblFichier
            // 
            lblFichier.AutoSize = true;
            lblFichier.Location = new Point(15, 9);
            lblFichier.Margin = new Padding(6, 0, 6, 0);
            lblFichier.Name = "lblFichier";
            lblFichier.Size = new Size(84, 32);
            lblFichier.TabIndex = 0;
            lblFichier.Text = "Fichier";
            // 
            // toolTip1
            // 
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 0;
            toolTip1.IsBalloon = true;
            toolTip1.ReshowDelay = 500;
            // 
            // pnlStatusBar
            // 
            pnlStatusBar.Dock = DockStyle.Bottom;
            pnlStatusBar.Location = new Point(0, 369);
            pnlStatusBar.Margin = new Padding(4, 2, 4, 2);
            pnlStatusBar.MaximumSize = new Size(0, 9);
            pnlStatusBar.Name = "pnlStatusBar";
            pnlStatusBar.Size = new Size(509, 9);
            pnlStatusBar.TabIndex = 3;
            // 
            // lblInfo
            // 
            lblInfo.AutoSize = true;
            lblInfo.Location = new Point(27, 301);
            lblInfo.Margin = new Padding(4, 0, 4, 0);
            lblInfo.Name = "lblInfo";
            lblInfo.Size = new Size(453, 32);
            lblInfo.TabIndex = 4;
            lblInfo.Text = "Déposez ici un relevé de notes d'Apogée";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(13F, 32F);
            AutoScaleMode = AutoScaleMode.Font;
            BackgroundImage = Properties.Resources.aposplit4;
            BackgroundImageLayout = ImageLayout.Zoom;
            ClientSize = new Size(509, 378);
            Controls.Add(lblInfo);
            Controls.Add(pnlStatusBar);
            Controls.Add(lblFichier);
            DoubleBuffered = true;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(6);
            MaximizeBox = false;
            Name = "Form1";
            Text = "Aposplit by BM";
            Load += Form1_Load;
            DragDrop += Form1_DragDrop;
            DragEnter += Form1_DragEnter;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Panel pnlStatusBar; 
        private Label lblFichier;
        private ToolTip toolTip1;
        private Label lblInfo;
    }
}
