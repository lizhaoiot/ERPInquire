namespace ERPInquire.CustomControl
{
    partial class PDFHELP
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PDFHELP));
            this.pdfDocument1 = new O2S.Components.PDFView4NET.PDFDocument(this.components);
            this.pdfPageView1 = new O2S.Components.PDFView4NET.PDFPageView();
            this.SuspendLayout();
            // 
            // pdfDocument1
            // 
            this.pdfDocument1.Metadata = null;
            this.pdfDocument1.PageLayout = O2S.Components.PDFView4NET.PDFPageLayout.SinglePage;
            this.pdfDocument1.PageMode = O2S.Components.PDFView4NET.PDFPageMode.UseNone;
            // 
            // pdfPageView1
            // 
            this.pdfPageView1.AutoScroll = true;
            this.pdfPageView1.DefaultEllipseAnnotationBorderWidth = 1D;
            this.pdfPageView1.DefaultInkAnnotationWidth = 1D;
            this.pdfPageView1.DefaultRectangleAnnotationBorderWidth = 1D;
            this.pdfPageView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pdfPageView1.Document = this.pdfDocument1;
            this.pdfPageView1.DownscaleLargeImages = false;
            this.pdfPageView1.EnableRepeatedKeys = false;
            this.pdfPageView1.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.pdfPageView1.Location = new System.Drawing.Point(0, 0);
            this.pdfPageView1.Name = "pdfPageView1";
            this.pdfPageView1.PageDisplayLayout = O2S.Components.PDFView4NET.PDFPageDisplayLayout.OneColumn;
            this.pdfPageView1.PageNumber = 0;
            this.pdfPageView1.RenderingProgressColor = System.Drawing.Color.Empty;
            this.pdfPageView1.RequiredFormFieldHighlightColor = System.Drawing.Color.Empty;
            this.pdfPageView1.ScrollPosition = new System.Drawing.Point(0, 0);
            this.pdfPageView1.Size = new System.Drawing.Size(689, 500);
            this.pdfPageView1.SubstituteFonts = null;
            this.pdfPageView1.TabIndex = 0;
            this.pdfPageView1.WorkMode = O2S.Components.PDFView4NET.UserInteractiveWorkMode.None;
            this.pdfPageView1.ZoomMode = O2S.Components.PDFView4NET.PDFZoomMode.FitWidth;
            // 
            // PDFHELP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(689, 500);
            this.Controls.Add(this.pdfPageView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "PDFHELP";
            this.Text = "操作说明文档";
            this.Load += new System.EventHandler(this.PDFHELP_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private O2S.Components.PDFView4NET.PDFDocument pdfDocument1;
        private O2S.Components.PDFView4NET.PDFPageView pdfPageView1;
    }
}