namespace MagosaAddIn.UI
{
    partial class CustomRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CustomRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnDivideShape = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnAlignToLeft = this.Factory.CreateRibbonButton();
            this.btnAlignToRight = this.Factory.CreateRibbonButton();
            this.btnAlignToTop = this.Factory.CreateRibbonButton();
            this.btnAlignToBottom = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnAlignLeftToRight = this.Factory.CreateRibbonButton();
            this.btnAlignRightToLeft = this.Factory.CreateRibbonButton();
            this.btnAlignTopToBottom = this.Factory.CreateRibbonButton();
            this.btnAlignBottomToTop = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnAlignToCenterHorizontal = this.Factory.CreateRibbonButton();
            this.btnAlignToCenterVertical = this.Factory.CreateRibbonButton();
            this.btnDistributeHorizontal = this.Factory.CreateRibbonButton();
            this.btnDistributeVertical = this.Factory.CreateRibbonButton();
            this.btnArrangeInGrid = this.Factory.CreateRibbonButton();
            this.btnArrangeInCircle = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "Magosa Tools";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnDivideShape);
            this.group1.Label = "図形分割";
            this.group1.Name = "group1";
            // 
            // btnDivideShape
            // 
            this.btnDivideShape.Label = "図形分割";
            this.btnDivideShape.Name = "btnDivideShape";
            this.btnDivideShape.OfficeImageId = "TableInsertGallery";
            this.btnDivideShape.ShowImage = true;
            this.btnDivideShape.SuperTip = "選択した四角形を指定した行・列数で分割します。水平・垂直マージンを個別に設定できます。";
            this.btnDivideShape.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDivideShape_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnAlignToLeft);
            this.group2.Items.Add(this.btnAlignToRight);
            this.group2.Items.Add(this.btnAlignToTop);
            this.group2.Items.Add(this.btnAlignToBottom);
            this.group2.Label = "基準整列";
            this.group2.Name = "group2";
            // 
            // btnAlignToLeft
            // 
            this.btnAlignToLeft.Label = "左端揃え";
            this.btnAlignToLeft.Name = "btnAlignToLeft";
            this.btnAlignToLeft.OfficeImageId = "ObjectAlignLeft";
            this.btnAlignToLeft.ShowImage = true;
            this.btnAlignToLeft.SuperTip = "基準図形の左端に、その他の図形の左端を揃えます。";
            this.btnAlignToLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToLeft_Click);
            // 
            // btnAlignToRight
            // 
            this.btnAlignToRight.Label = "右端揃え";
            this.btnAlignToRight.Name = "btnAlignToRight";
            this.btnAlignToRight.OfficeImageId = "ObjectAlignRight";
            this.btnAlignToRight.ShowImage = true;
            this.btnAlignToRight.SuperTip = "基準図形の右端に、その他の図形の右端を揃えます。";
            this.btnAlignToRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToRight_Click);
            // 
            // btnAlignToTop
            // 
            this.btnAlignToTop.Label = "上端揃え";
            this.btnAlignToTop.Name = "btnAlignToTop";
            this.btnAlignToTop.OfficeImageId = "ObjectAlignTop";
            this.btnAlignToTop.ShowImage = true;
            this.btnAlignToTop.SuperTip = "基準図形の上端に、その他の図形の上端を揃えます。";
            this.btnAlignToTop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToTop_Click);
            // 
            // btnAlignToBottom
            // 
            this.btnAlignToBottom.Label = "下端揃え";
            this.btnAlignToBottom.Name = "btnAlignToBottom";
            this.btnAlignToBottom.OfficeImageId = "ObjectAlignBottom";
            this.btnAlignToBottom.ShowImage = true;
            this.btnAlignToBottom.SuperTip = "基準図形の下端に、その他の図形の下端を揃えます。";
            this.btnAlignToBottom.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToBottom_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnAlignLeftToRight);
            this.group3.Items.Add(this.btnAlignRightToLeft);
            this.group3.Items.Add(this.btnAlignTopToBottom);
            this.group3.Items.Add(this.btnAlignBottomToTop);
            this.group3.Label = "隣接整列";
            this.group3.Name = "group3";
            // 
            // btnAlignLeftToRight
            // 
            this.btnAlignLeftToRight.Label = "左端→右端";
            this.btnAlignLeftToRight.Name = "btnAlignLeftToRight";
            this.btnAlignLeftToRight.OfficeImageId = "AlignLeft";
            this.btnAlignLeftToRight.ShowImage = true;
            this.btnAlignLeftToRight.SuperTip = "基準図形の左端に、その他の図形の右端を隣接させます。";
            this.btnAlignLeftToRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignLeftToRight_Click);
            // 
            // btnAlignRightToLeft
            // 
            this.btnAlignRightToLeft.Label = "右端→左端";
            this.btnAlignRightToLeft.Name = "btnAlignRightToLeft";
            this.btnAlignRightToLeft.OfficeImageId = "AlignRight";
            this.btnAlignRightToLeft.ShowImage = true;
            this.btnAlignRightToLeft.SuperTip = "基準図形の右端に、その他の図形の左端を隣接させます。";
            this.btnAlignRightToLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignRightToLeft_Click);
            // 
            // btnAlignTopToBottom
            // 
            this.btnAlignTopToBottom.Label = "上端→下端";
            this.btnAlignTopToBottom.Name = "btnAlignTopToBottom";
            this.btnAlignTopToBottom.OfficeImageId = "AlignTop";
            this.btnAlignTopToBottom.ShowImage = true;
            this.btnAlignTopToBottom.SuperTip = "基準図形の上端に、その他の図形の下端を隣接させます。";
            this.btnAlignTopToBottom.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignTopToBottom_Click);
            // 
            // btnAlignBottomToTop
            // 
            this.btnAlignBottomToTop.Label = "下端→上端";
            this.btnAlignBottomToTop.Name = "btnAlignBottomToTop";
            this.btnAlignBottomToTop.OfficeImageId = "AlignBottom";
            this.btnAlignBottomToTop.ShowImage = true;
            this.btnAlignBottomToTop.SuperTip = "基準図形の下端に、その他の図形の上端を隣接させます。";
            this.btnAlignBottomToTop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignBottomToTop_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnAlignToCenterHorizontal);
            this.group4.Items.Add(this.btnAlignToCenterVertical);
            this.group4.Items.Add(this.btnDistributeHorizontal);
            this.group4.Items.Add(this.btnDistributeVertical);
            this.group4.Items.Add(this.btnArrangeInGrid);
            this.group4.Items.Add(this.btnArrangeInCircle);
            this.group4.Label = "拡張整列";
            this.group4.Name = "group4";
            // 
            // btnAlignToCenterHorizontal
            // 
            this.btnAlignToCenterHorizontal.Label = "水平中央";
            this.btnAlignToCenterHorizontal.Name = "btnAlignToCenterHorizontal";
            this.btnAlignToCenterHorizontal.OfficeImageId = "ObjectAlignCenterHorizontal";
            this.btnAlignToCenterHorizontal.ShowImage = true;
            this.btnAlignToCenterHorizontal.SuperTip = "基準図形の水平中央に、その他の図形の中央を揃えます。";
            this.btnAlignToCenterHorizontal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToCenterHorizontal_Click);
            // 
            // btnAlignToCenterVertical
            // 
            this.btnAlignToCenterVertical.Label = "垂直中央";
            this.btnAlignToCenterVertical.Name = "btnAlignToCenterVertical";
            this.btnAlignToCenterVertical.OfficeImageId = "ObjectAlignCenterVertical";
            this.btnAlignToCenterVertical.ShowImage = true;
            this.btnAlignToCenterVertical.SuperTip = "基準図形の垂直中央に、その他の図形の中央を揃えます。";
            this.btnAlignToCenterVertical.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToCenterVertical_Click);
            // 
            // btnDistributeHorizontal
            // 
            this.btnDistributeHorizontal.Label = "水平等間隔";
            this.btnDistributeHorizontal.Name = "btnDistributeHorizontal";
            this.btnDistributeHorizontal.OfficeImageId = "ObjectsDistributeHorizontally";
            this.btnDistributeHorizontal.ShowImage = true;
            this.btnDistributeHorizontal.SuperTip = "選択した図形を水平方向に等間隔で配置します。（3つ以上の図形が必要）";
            this.btnDistributeHorizontal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDistributeHorizontal_Click);
            // 
            // btnDistributeVertical
            // 
            this.btnDistributeVertical.Label = "垂直等間隔";
            this.btnDistributeVertical.Name = "btnDistributeVertical";
            this.btnDistributeVertical.OfficeImageId = "ObjectsDistributeVertically";
            this.btnDistributeVertical.ShowImage = true;
            this.btnDistributeVertical.SuperTip = "選択した図形を垂直方向に等間隔で配置します。（3つ以上の図形が必要）";
            this.btnDistributeVertical.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDistributeVertical_Click);
            // 
            // btnArrangeInGrid
            // 
            this.btnArrangeInGrid.Label = "グリッド配置";
            this.btnArrangeInGrid.Name = "btnArrangeInGrid";
            this.btnArrangeInGrid.OfficeImageId = "ViewGridlines";
            this.btnArrangeInGrid.ShowImage = true;
            this.btnArrangeInGrid.SuperTip = "選択した図形を指定した列数のグリッド状に配置します。";
            this.btnArrangeInGrid.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnArrangeInGrid_Click);
            // 
            // btnArrangeInCircle
            // 
            this.btnArrangeInCircle.Label = "円形配置";
            this.btnArrangeInCircle.Name = "btnArrangeInCircle";
            this.btnArrangeInCircle.OfficeImageId = "ShapeOval";
            this.btnArrangeInCircle.ShowImage = true;
            this.btnArrangeInCircle.SuperTip = "選択した図形を指定した中心と半径で円形に配置します。";
            this.btnArrangeInCircle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnArrangeInCircle_Click);
            // 
            // CustomRibbon
            // 
            this.Name = "CustomRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CustomRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;

        // Group 1: 図形分割
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDivideShape;

        // Group 2: 基準整列 (順番変更)
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToBottom;

        // Group 3: 隣接整列 (名前変更・順番変更)
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignLeftToRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignRightToLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignTopToBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignBottomToTop;

        // Group 4: 拡張整列
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToCenterHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToCenterVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDistributeHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDistributeVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnArrangeInGrid;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnArrangeInCircle;
    }

    partial class ThisRibbonCollection
    {
        //internal CustomRibbon CustomRibbon
        //{
        //    get { return this.GetRibbon<CustomRibbon>(); }
        //}
    }
}
