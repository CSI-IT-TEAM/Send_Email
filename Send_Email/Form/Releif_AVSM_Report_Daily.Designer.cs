
namespace Send_Email
{
    partial class Releif_AVSM_Report_Daily
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.grdMain2 = new JPlatform.Client.Controls6.GridControlEx();
            this.grdView2 = new JPlatform.Client.Controls6.BandedGridViewEx();
            this.PLANT_CD = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.PLANT_NM = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.LINE_CD = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.AREA_NM = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.PROCESS_CD = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.TO_QTY = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.PO_WS = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.PO_RELIEF = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.PO_MAT_HANDLER = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.PO_OTHER_LINE = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.PO_TOTAL = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.BALANCE = new JPlatform.Client.Controls6.BandedGridColumnEx();
            this.gridBand1 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand2 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand3 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand4 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand5 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand6 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand7 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand9 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand8 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand11 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand10 = new JPlatform.Client.Controls6.GridBandEx();
            this.gridBand12 = new JPlatform.Client.Controls6.GridBandEx();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdMain2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdView2)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.grdMain2);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1031, 212);
            this.panel1.TabIndex = 0;
            // 
            // grdMain2
            // 
            this.grdMain2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grdMain2.Location = new System.Drawing.Point(0, 0);
            this.grdMain2.MainView = this.grdView2;
            this.grdMain2.Name = "grdMain2";
            this.grdMain2.Size = new System.Drawing.Size(1031, 212);
            this.grdMain2.TabIndex = 240;
            this.grdMain2.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grdView2});
            // 
            // grdView2
            // 
            this.grdView2.ActionMode = JPlatform.Client.Controls6.ActionMode.View;
            this.grdView2.Appearance.FooterPanel.Options.UseTextOptions = true;
            this.grdView2.Appearance.FooterPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.grdView2.Appearance.Row.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grdView2.Appearance.Row.Options.UseFont = true;
            this.grdView2.Appearance.Row.Options.UseTextOptions = true;
            this.grdView2.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.grdView2.Appearance.Row.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.grdView2.BandPanelRowHeight = 30;
            this.grdView2.Bands.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] {
            this.gridBand1,
            this.gridBand2,
            this.gridBand3,
            this.gridBand4,
            this.gridBand5,
            this.gridBand6,
            this.gridBand12});
            this.grdView2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.grdView2.Columns.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn[] {
            this.PLANT_CD,
            this.PLANT_NM,
            this.LINE_CD,
            this.AREA_NM,
            this.PROCESS_CD,
            this.TO_QTY,
            this.PO_WS,
            this.PO_RELIEF,
            this.PO_MAT_HANDLER,
            this.PO_OTHER_LINE,
            this.PO_TOTAL,
            this.BALANCE});
            this.grdView2.GridControl = this.grdMain2;
            this.grdView2.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Never;
            this.grdView2.Name = "grdView2";
            this.grdView2.OptionsCustomization.AllowBandMoving = false;
            this.grdView2.OptionsCustomization.AllowColumnMoving = false;
            this.grdView2.OptionsCustomization.AllowGroup = false;
            this.grdView2.OptionsCustomization.AllowSort = false;
            this.grdView2.OptionsPrint.PrintHeader = false;
            this.grdView2.OptionsSelection.CheckBoxSelectorColumnWidth = 25;
            this.grdView2.OptionsSelection.MultiSelect = true;
            this.grdView2.OptionsView.AllowCellMerge = true;
            this.grdView2.OptionsView.ColumnAutoWidth = false;
            this.grdView2.OptionsView.ShowColumnHeaders = false;
            this.grdView2.OptionsView.ShowGroupPanel = false;
            this.grdView2.OptionsView.ShowIndicator = false;
            this.grdView2.RowHeight = 30;
            this.grdView2.SaveSPName = null;
            this.grdView2.VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Never;
            this.grdView2.ViewSPName = null;
            // 
            // PLANT_CD
            // 
            this.PLANT_CD.BindingField = "PLANT_CD";
            this.PLANT_CD.Caption = "PLANT_CD";
            this.PLANT_CD.ColumnEdit = null;
            this.PLANT_CD.FieldName = "PLANT_CD";
            this.PLANT_CD.Name = "PLANT_CD";
            this.PLANT_CD.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.PLANT_CD.Visible = true;
            // 
            // PLANT_NM
            // 
            this.PLANT_NM.BindingField = "PLANT_NM";
            this.PLANT_NM.Caption = "PLANT_NM";
            this.PLANT_NM.ColumnEdit = null;
            this.PLANT_NM.FieldName = "PLANT_NM";
            this.PLANT_NM.Name = "PLANT_NM";
            this.PLANT_NM.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.PLANT_NM.Visible = true;
            // 
            // LINE_CD
            // 
            this.LINE_CD.Caption = "LINE_CD";
            this.LINE_CD.ColumnEdit = null;
            this.LINE_CD.Name = "LINE_CD";
            this.LINE_CD.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.LINE_CD.Visible = true;
            // 
            // AREA_NM
            // 
            this.AREA_NM.BindingField = "AREA_NM";
            this.AREA_NM.Caption = "AREA_NM";
            this.AREA_NM.ColumnEdit = null;
            this.AREA_NM.FieldName = "AREA_NM";
            this.AREA_NM.Name = "AREA_NM";
            this.AREA_NM.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.AREA_NM.Visible = true;
            // 
            // PROCESS_CD
            // 
            this.PROCESS_CD.BindingField = "PROCESS_CD";
            this.PROCESS_CD.Caption = "PROCESS_CD";
            this.PROCESS_CD.ColumnEdit = null;
            this.PROCESS_CD.FieldName = "PROCESS_CD";
            this.PROCESS_CD.Name = "PROCESS_CD";
            this.PROCESS_CD.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.PROCESS_CD.Visible = true;
            // 
            // TO_QTY
            // 
            this.TO_QTY.BindingField = "TO_QTY";
            this.TO_QTY.Caption = "TO_QTY";
            this.TO_QTY.ColumnEdit = null;
            this.TO_QTY.FieldName = "TO_QTY";
            this.TO_QTY.Name = "TO_QTY";
            this.TO_QTY.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.TO_QTY.Visible = true;
            // 
            // PO_WS
            // 
            this.PO_WS.BindingField = "PO_WS";
            this.PO_WS.Caption = "PO_WS";
            this.PO_WS.ColumnEdit = null;
            this.PO_WS.FieldName = "PO_WS";
            this.PO_WS.Name = "PO_WS";
            this.PO_WS.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.PO_WS.Visible = true;
            this.PO_WS.Width = 100;
            // 
            // PO_RELIEF
            // 
            this.PO_RELIEF.BindingField = "PO_RELIEF";
            this.PO_RELIEF.Caption = "PO_RELIEF";
            this.PO_RELIEF.ColumnEdit = null;
            this.PO_RELIEF.FieldName = "PO_RELIEF";
            this.PO_RELIEF.Name = "PO_RELIEF";
            this.PO_RELIEF.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.PO_RELIEF.Visible = true;
            this.PO_RELIEF.Width = 100;
            // 
            // PO_MAT_HANDLER
            // 
            this.PO_MAT_HANDLER.BindingField = "PO_MAT_HANDLER";
            this.PO_MAT_HANDLER.Caption = "PO_MAT_HANDLER";
            this.PO_MAT_HANDLER.ColumnEdit = null;
            this.PO_MAT_HANDLER.FieldName = "PO_MAT_HANDLER";
            this.PO_MAT_HANDLER.Name = "PO_MAT_HANDLER";
            this.PO_MAT_HANDLER.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.PO_MAT_HANDLER.Visible = true;
            this.PO_MAT_HANDLER.Width = 100;
            // 
            // PO_OTHER_LINE
            // 
            this.PO_OTHER_LINE.BindingField = "PO_OTHER_LINE";
            this.PO_OTHER_LINE.Caption = "PO_OTHER_LINE";
            this.PO_OTHER_LINE.ColumnEdit = null;
            this.PO_OTHER_LINE.FieldName = "PO_OTHER_LINE";
            this.PO_OTHER_LINE.Name = "PO_OTHER_LINE";
            this.PO_OTHER_LINE.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.PO_OTHER_LINE.Visible = true;
            this.PO_OTHER_LINE.Width = 100;
            // 
            // PO_TOTAL
            // 
            this.PO_TOTAL.BindingField = "PO_TOTAL";
            this.PO_TOTAL.Caption = "PO_TOTAL";
            this.PO_TOTAL.ColumnEdit = null;
            this.PO_TOTAL.FieldName = "PO_TOTAL";
            this.PO_TOTAL.Name = "PO_TOTAL";
            this.PO_TOTAL.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.PO_TOTAL.Visible = true;
            this.PO_TOTAL.Width = 100;
            // 
            // BALANCE
            // 
            this.BALANCE.BindingField = "BALANCE";
            this.BALANCE.Caption = "BALANCE";
            this.BALANCE.ColumnEdit = null;
            this.BALANCE.FieldName = "BALANCE";
            this.BALANCE.Name = "BALANCE";
            this.BALANCE.SortMode = DevExpress.XtraGrid.ColumnSortMode.Default;
            this.BALANCE.Visible = true;
            // 
            // gridBand1
            // 
            this.gridBand1.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand1.AppearanceHeader.Options.UseFont = true;
            this.gridBand1.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand1.Caption = "Plant";
            this.gridBand1.Columns.Add(this.PLANT_CD);
            this.gridBand1.Columns.Add(this.PLANT_NM);
            this.gridBand1.Name = "gridBand1";
            this.gridBand1.VisibleIndex = 0;
            this.gridBand1.Width = 150;
            // 
            // gridBand2
            // 
            this.gridBand2.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand2.AppearanceHeader.Options.UseFont = true;
            this.gridBand2.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand2.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand2.Caption = "Line";
            this.gridBand2.Columns.Add(this.LINE_CD);
            this.gridBand2.Name = "gridBand2";
            this.gridBand2.VisibleIndex = 1;
            this.gridBand2.Width = 75;
            // 
            // gridBand3
            // 
            this.gridBand3.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand3.AppearanceHeader.Options.UseFont = true;
            this.gridBand3.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand3.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand3.Caption = "Area";
            this.gridBand3.Columns.Add(this.AREA_NM);
            this.gridBand3.Name = "gridBand3";
            this.gridBand3.VisibleIndex = 2;
            this.gridBand3.Width = 75;
            // 
            // gridBand4
            // 
            this.gridBand4.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand4.AppearanceHeader.Options.UseFont = true;
            this.gridBand4.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand4.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand4.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand4.Caption = "Process";
            this.gridBand4.Columns.Add(this.PROCESS_CD);
            this.gridBand4.Name = "gridBand4";
            this.gridBand4.VisibleIndex = 3;
            this.gridBand4.Width = 75;
            // 
            // gridBand5
            // 
            this.gridBand5.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand5.AppearanceHeader.Options.UseFont = true;
            this.gridBand5.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand5.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand5.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand5.Caption = "TO";
            this.gridBand5.Columns.Add(this.TO_QTY);
            this.gridBand5.Name = "gridBand5";
            this.gridBand5.VisibleIndex = 4;
            this.gridBand5.Width = 75;
            // 
            // gridBand6
            // 
            this.gridBand6.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand6.AppearanceHeader.Options.UseFont = true;
            this.gridBand6.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand6.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand6.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand6.Caption = "PO";
            this.gridBand6.Children.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] {
            this.gridBand7,
            this.gridBand9,
            this.gridBand8,
            this.gridBand11,
            this.gridBand10});
            this.gridBand6.Name = "gridBand6";
            this.gridBand6.VisibleIndex = 5;
            this.gridBand6.Width = 500;
            // 
            // gridBand7
            // 
            this.gridBand7.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand7.AppearanceHeader.Options.UseFont = true;
            this.gridBand7.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand7.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand7.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand7.Caption = "Workshop";
            this.gridBand7.Columns.Add(this.PO_WS);
            this.gridBand7.Name = "gridBand7";
            this.gridBand7.VisibleIndex = 0;
            this.gridBand7.Width = 100;
            // 
            // gridBand9
            // 
            this.gridBand9.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand9.AppearanceHeader.Options.UseFont = true;
            this.gridBand9.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand9.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand9.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand9.Caption = "Relief";
            this.gridBand9.Columns.Add(this.PO_RELIEF);
            this.gridBand9.Name = "gridBand9";
            this.gridBand9.VisibleIndex = 1;
            this.gridBand9.Width = 100;
            // 
            // gridBand8
            // 
            this.gridBand8.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand8.AppearanceHeader.Options.UseFont = true;
            this.gridBand8.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand8.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand8.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand8.Caption = "Other line";
            this.gridBand8.Columns.Add(this.PO_OTHER_LINE);
            this.gridBand8.Name = "gridBand8";
            this.gridBand8.VisibleIndex = 2;
            this.gridBand8.Width = 100;
            // 
            // gridBand11
            // 
            this.gridBand11.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand11.AppearanceHeader.Options.UseFont = true;
            this.gridBand11.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand11.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand11.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand11.Caption = "Material handler";
            this.gridBand11.Columns.Add(this.PO_MAT_HANDLER);
            this.gridBand11.Name = "gridBand11";
            this.gridBand11.VisibleIndex = 3;
            this.gridBand11.Width = 100;
            // 
            // gridBand10
            // 
            this.gridBand10.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridBand10.AppearanceHeader.Options.UseFont = true;
            this.gridBand10.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand10.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand10.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand10.Caption = "Total";
            this.gridBand10.Columns.Add(this.PO_TOTAL);
            this.gridBand10.Name = "gridBand10";
            this.gridBand10.VisibleIndex = 4;
            this.gridBand10.Width = 100;
            // 
            // gridBand12
            // 
            this.gridBand12.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.gridBand12.AppearanceHeader.Options.UseFont = true;
            this.gridBand12.AppearanceHeader.Options.UseTextOptions = true;
            this.gridBand12.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridBand12.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.gridBand12.Caption = "Balance";
            this.gridBand12.Columns.Add(this.BALANCE);
            this.gridBand12.Name = "gridBand12";
            this.gridBand12.VisibleIndex = 6;
            this.gridBand12.Width = 75;
            // 
            // Releif_AVSM_Report_Daily
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1146, 408);
            this.Controls.Add(this.panel1);
            this.Name = "Releif_AVSM_Report_Daily";
            this.Text = "Releif_AVSM_Report_Daily";
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grdMain2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdView2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private JPlatform.Client.Controls6.GridControlEx grdMain2;
        private JPlatform.Client.Controls6.BandedGridViewEx grdView2;
        private JPlatform.Client.Controls6.GridBandEx gridBand1;
        private JPlatform.Client.Controls6.BandedGridColumnEx PLANT_CD;
        private JPlatform.Client.Controls6.BandedGridColumnEx PLANT_NM;
        private JPlatform.Client.Controls6.GridBandEx gridBand2;
        private JPlatform.Client.Controls6.BandedGridColumnEx LINE_CD;
        private JPlatform.Client.Controls6.GridBandEx gridBand3;
        private JPlatform.Client.Controls6.BandedGridColumnEx AREA_NM;
        private JPlatform.Client.Controls6.GridBandEx gridBand4;
        private JPlatform.Client.Controls6.BandedGridColumnEx PROCESS_CD;
        private JPlatform.Client.Controls6.GridBandEx gridBand5;
        private JPlatform.Client.Controls6.BandedGridColumnEx TO_QTY;
        private JPlatform.Client.Controls6.GridBandEx gridBand6;
        private JPlatform.Client.Controls6.GridBandEx gridBand7;
        private JPlatform.Client.Controls6.BandedGridColumnEx PO_WS;
        private JPlatform.Client.Controls6.GridBandEx gridBand9;
        private JPlatform.Client.Controls6.BandedGridColumnEx PO_RELIEF;
        private JPlatform.Client.Controls6.GridBandEx gridBand8;
        private JPlatform.Client.Controls6.BandedGridColumnEx PO_OTHER_LINE;
        private JPlatform.Client.Controls6.GridBandEx gridBand11;
        private JPlatform.Client.Controls6.BandedGridColumnEx PO_MAT_HANDLER;
        private JPlatform.Client.Controls6.GridBandEx gridBand10;
        private JPlatform.Client.Controls6.BandedGridColumnEx PO_TOTAL;
        private JPlatform.Client.Controls6.GridBandEx gridBand12;
        private JPlatform.Client.Controls6.BandedGridColumnEx BALANCE;
    }
}