namespace DWGViewer
{
    partial class DWGViewer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DWGViewer));
            this.ACADViewer = new AxACCTRLLib.AxAcCtrl();
            ((System.ComponentModel.ISupportInitialize)(this.ACADViewer)).BeginInit();
            this.SuspendLayout();
            // 
            // ACADViewer
            // 
            this.ACADViewer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ACADViewer.Enabled = true;
            this.ACADViewer.Location = new System.Drawing.Point(0, 0);
            this.ACADViewer.Name = "ACADViewer";
            this.ACADViewer.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("ACADViewer.OcxState")));
            this.ACADViewer.Size = new System.Drawing.Size(491, 417);
            this.ACADViewer.TabIndex = 0;
            // 
            // DWGViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.ACADViewer);
            this.Name = "DWGViewer";
            this.Size = new System.Drawing.Size(491, 417);
            ((System.ComponentModel.ISupportInitialize)(this.ACADViewer)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AxACCTRLLib.AxAcCtrl ACADViewer;
    }
}
