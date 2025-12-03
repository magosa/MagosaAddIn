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
            this.btnAlignAndDistributeHorizontal = this.Factory.CreateRibbonButton();
            this.btnAlignAndDistributeVertical = this.Factory.CreateRibbonButton();
            this.btnArrangeInGrid = this.Factory.CreateRibbonButton();
            this.btnArrangeHorizontalWithMargin = this.Factory.CreateRibbonButton();
            this.btnArrangeVerticalWithMargin = this.Factory.CreateRibbonButton();
            this.btnArrangeInCircle = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.btnSelectSameFormat = this.Factory.CreateRibbonButton();
            // 新規追加: ハンドル調整グループ
            this.group6 = this.Factory.CreateRibbonGroup();
            this.btnAdjustmentHandles = this.Factory.CreateRibbonButton();
            this.btnAngleHandles = this.Factory.CreateRibbonButton();
            this.btnResetAdjustments = this.Factory.CreateRibbonButton();

            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.group6.SuspendLayout(); // 新規追加
            this.SuspendLayout();

            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group6); // 新規追加
            this.tab1.Label = "Magosa Tools";
            this.tab1.Name = "tab1";

            // 
            // group1
            // 
            this.group1.Items.Add(this.btnDivideShape);
            this.group1.Label = "図形操作";
            this.group1.Name = "group1";

            // 
            // btnDivideShape
            // 
            this.btnDivideShape.Label = "グリッド分割";
            this.btnDivideShape.Name = "btnDivideShape";
            this.btnDivideShape.OfficeImageId = "AppointmentColorDialog";
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
            this.btnAlignToLeft.OfficeImageId = "ObjectsAlignLeftSmart";
            this.btnAlignToLeft.ShowImage = true;
            this.btnAlignToLeft.SuperTip = "基準図形の左端に、その他の図形の左端を揃えます。";
            this.btnAlignToLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToLeft_Click);

            // 
            // btnAlignToRight
            // 
            this.btnAlignToRight.Label = "右端揃え";
            this.btnAlignToRight.Name = "btnAlignToRight";
            this.btnAlignToRight.OfficeImageId = "ObjectsAlignRightSmart";
            this.btnAlignToRight.ShowImage = true;
            this.btnAlignToRight.SuperTip = "基準図形の右端に、その他の図形の右端を揃えます。";
            this.btnAlignToRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToRight_Click);

            // 
            // btnAlignToTop
            // 
            this.btnAlignToTop.Label = "上端揃え";
            this.btnAlignToTop.Name = "btnAlignToTop";
            this.btnAlignToTop.OfficeImageId = "ObjectsAlignTopSmart";
            this.btnAlignToTop.ShowImage = true;
            this.btnAlignToTop.SuperTip = "基準図形の上端に、その他の図形の上端を揃えます。";
            this.btnAlignToTop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToTop_Click);

            // 
            // btnAlignToBottom
            // 
            this.btnAlignToBottom.Label = "下端揃え";
            this.btnAlignToBottom.Name = "btnAlignToBottom";
            this.btnAlignToBottom.OfficeImageId = "ObjectsAlignBottomSmart";
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
            this.btnAlignLeftToRight.Label = "左端に右端を隣接";
            this.btnAlignLeftToRight.Name = "btnAlignLeftToRight";
            this.btnAlignLeftToRight.OfficeImageId = "SnapToAlignmentBox";
            this.btnAlignLeftToRight.ShowImage = true;
            this.btnAlignLeftToRight.SuperTip = "基準図形の左端に、その他の図形の右端を隣接させます。";
            this.btnAlignLeftToRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignLeftToRight_Click);

            // 
            // btnAlignRightToLeft
            // 
            this.btnAlignRightToLeft.Label = "右端に左端を隣接";
            this.btnAlignRightToLeft.Name = "btnAlignRightToLeft";
            this.btnAlignRightToLeft.OfficeImageId = "SnapToAlignmentBox";
            this.btnAlignRightToLeft.ShowImage = true;
            this.btnAlignRightToLeft.SuperTip = "基準図形の右端に、その他の図形の左端を隣接させます。";
            this.btnAlignRightToLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignRightToLeft_Click);

            // 
            // btnAlignTopToBottom
            // 
            this.btnAlignTopToBottom.Label = "上端に下端を隣接";
            this.btnAlignTopToBottom.Name = "btnAlignTopToBottom";
            this.btnAlignTopToBottom.OfficeImageId = "SnapToAlignmentBox";
            this.btnAlignTopToBottom.ShowImage = true;
            this.btnAlignTopToBottom.SuperTip = "基準図形の上端に、その他の図形の下端を隣接させます。";
            this.btnAlignTopToBottom.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignTopToBottom_Click);

            // 
            // btnAlignBottomToTop
            // 
            this.btnAlignBottomToTop.Label = "下端に上端を隣接";
            this.btnAlignBottomToTop.Name = "btnAlignBottomToTop";
            this.btnAlignBottomToTop.OfficeImageId = "SnapToAlignmentBox";
            this.btnAlignBottomToTop.ShowImage = true;
            this.btnAlignBottomToTop.SuperTip = "基準図形の下端に、その他の図形の上端を隣接させます。";
            this.btnAlignBottomToTop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignBottomToTop_Click);

            // 
            // group4
            // 
            this.group4.Items.Add(this.btnAlignAndDistributeHorizontal);
            this.group4.Items.Add(this.btnAlignAndDistributeVertical);
            this.group4.Items.Add(this.btnArrangeInGrid);
            this.group4.Items.Add(this.btnArrangeHorizontalWithMargin);
            this.group4.Items.Add(this.btnArrangeVerticalWithMargin);
            this.group4.Items.Add(this.btnArrangeInCircle);
            this.group4.Label = "拡張整列";
            this.group4.Name = "group4";

            // 
            // btnAlignAndDistributeHorizontal
            // 
            this.btnAlignAndDistributeHorizontal.Label = "水平中央・等間隔";
            this.btnAlignAndDistributeHorizontal.Name = "btnAlignAndDistributeHorizontal";
            this.btnAlignAndDistributeHorizontal.OfficeImageId = "AlignDistributeHorizontally";
            this.btnAlignAndDistributeHorizontal.ShowImage = true;
            this.btnAlignAndDistributeHorizontal.SuperTip = "基準図形の水平中央に揃えて等間隔で配置します。2つの図形の場合は中央揃えのみ実行されます。";
            this.btnAlignAndDistributeHorizontal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignAndDistributeHorizontal_Click);

            // 
            // btnAlignAndDistributeVertical
            // 
            this.btnAlignAndDistributeVertical.Label = "垂直中央・等間隔";
            this.btnAlignAndDistributeVertical.Name = "btnAlignAndDistributeVertical";
            this.btnAlignAndDistributeVertical.OfficeImageId = "AlignDistributeHorizontally";
            this.btnAlignAndDistributeVertical.ShowImage = true;
            this.btnAlignAndDistributeVertical.SuperTip = "基準図形の垂直中央に揃えて等間隔で配置します。2つの図形の場合は中央揃えのみ実行されます。";
            this.btnAlignAndDistributeVertical.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignAndDistributeVertical_Click);

            // 
            // btnArrangeInGrid
            // 
            this.btnArrangeInGrid.Label = "グリッド配置";
            this.btnArrangeInGrid.Name = "btnArrangeInGrid";
            this.btnArrangeInGrid.OfficeImageId = "SmartAlign";
            this.btnArrangeInGrid.ShowImage = true;
            this.btnArrangeInGrid.SuperTip = "選択した図形を指定した列数のグリッド状に配置します。水平・垂直間隔を個別に設定できます。";
            this.btnArrangeInGrid.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnArrangeInGrid_Click);

            // 
            // btnArrangeHorizontalWithMargin
            // 
            this.btnArrangeHorizontalWithMargin.Label = "水平マージン";
            this.btnArrangeHorizontalWithMargin.Name = "btnArrangeHorizontalWithMargin";
            this.btnArrangeHorizontalWithMargin.OfficeImageId = "ObjectsAlignDistributeHorizontallyRemove";
            this.btnArrangeHorizontalWithMargin.ShowImage = true;
            this.btnArrangeHorizontalWithMargin.SuperTip = "基準図形を中心に任意のマージンで水平方向に配置します。元の位置関係を保持します。";
            this.btnArrangeHorizontalWithMargin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnArrangeHorizontalWithMargin_Click);

            // 
            // btnArrangeVerticalWithMargin
            // 
            this.btnArrangeVerticalWithMargin.Label = "垂直マージン";
            this.btnArrangeVerticalWithMargin.Name = "btnArrangeVerticalWithMargin";
            this.btnArrangeVerticalWithMargin.OfficeImageId = "ObjectsAlignDistributeVerticallyRemove";
            this.btnArrangeVerticalWithMargin.ShowImage = true;
            this.btnArrangeVerticalWithMargin.SuperTip = "基準図形を中心に任意のマージンで垂直方向に配置します。元の位置関係を保持します。";
            this.btnArrangeVerticalWithMargin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnArrangeVerticalWithMargin_Click);

            // 
            // btnArrangeInCircle
            // 
            this.btnArrangeInCircle.Label = "円形配置";
            this.btnArrangeInCircle.Name = "btnArrangeInCircle";
            this.btnArrangeInCircle.OfficeImageId = "DiagramRadialInsertClassic";
            this.btnArrangeInCircle.ShowImage = true;
            this.btnArrangeInCircle.SuperTip = "選択した図形を指定した中心と半径で円形に配置します。選択図形の中心を自動取得することも可能です。";
            this.btnArrangeInCircle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnArrangeInCircle_Click);

            // 
            // group5
            // 
            this.group5.Items.Add(this.btnSelectSameFormat);
            this.group5.Label = "選択補助";
            this.group5.Name = "group5";

            // 
            // btnSelectSameFormat
            // 
            this.btnSelectSameFormat.Label = "同一書式選択";
            this.btnSelectSameFormat.Name = "btnSelectSameFormat";
            this.btnSelectSameFormat.OfficeImageId = "IconSelectArea";
            this.btnSelectSameFormat.ShowImage = true;
            this.btnSelectSameFormat.SuperTip = "選択中の図形を基準に、同じ書式を持つ図形を選択します。塗りつぶし色、枠線スタイル、またはその両方から選択条件を指定できます。";
            this.btnSelectSameFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectSameFormat_Click);

            // 
            // group6
            // 
            this.group6.Items.Add(this.btnAdjustmentHandles);
            this.group6.Items.Add(this.btnAngleHandles);
            this.group6.Items.Add(this.btnResetAdjustments);
            this.group6.Label = "ハンドル調整";
            this.group6.Name = "group6";

            // 
            // btnAdjustmentHandles
            // 
            this.btnAdjustmentHandles.Label = "調整ハンドル";
            this.btnAdjustmentHandles.Name = "btnAdjustmentHandles";
            this.btnAdjustmentHandles.OfficeImageId = "SelectCurrentRegion";
            this.btnAdjustmentHandles.ShowImage = true;
            this.btnAdjustmentHandles.SuperTip = "選択した図形の調整ハンドルを数値で精密に設定します。角丸四角形の角丸、吹き出しの尻尾位置、矢印の矢じりなどを調整できます。";
            this.btnAdjustmentHandles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAdjustmentHandles_Click);

            // 
            // btnAngleHandles
            // 
            this.btnAngleHandles.Label = "角度ハンドル";
            this.btnAngleHandles.Name = "btnAngleHandles";
            this.btnAngleHandles.OfficeImageId = "ObjectRotateFree";
            this.btnAngleHandles.ShowImage = true;
            this.btnAngleHandles.SuperTip = "円弧・弦・扇形・ブロック円弧・ドーナツ・三日月などの角度ハンドルを数値で精密に設定します。開始角度、終了角度、内径比率などを制御できます。";
            this.btnAngleHandles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAngleHandles_Click);

            // 
            // btnResetAdjustments
            // 
            this.btnResetAdjustments.Label = "リセット";
            this.btnResetAdjustments.Name = "btnResetAdjustments";
            this.btnResetAdjustments.OfficeImageId = "Undo";
            this.btnResetAdjustments.ShowImage = true;
            this.btnResetAdjustments.SuperTip = "選択した図形の調整ハンドルと角度ハンドルをデフォルト値にリセットします。";
            this.btnResetAdjustments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResetAdjustments_Click);

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
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group6.ResumeLayout(false); // 新規追加
            this.group6.PerformLayout(); // 新規追加
            this.ResumeLayout(false);
        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;

        // Group 1: 図形分割
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDivideShape;

        // Group 2: 基準整列
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToBottom;

        // Group 3: 隣接整列
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignLeftToRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignRightToLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignTopToBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignBottomToTop;

        // Group 4: 拡張整列
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignAndDistributeHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignAndDistributeVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnArrangeHorizontalWithMargin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnArrangeVerticalWithMargin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnArrangeInGrid;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnArrangeInCircle;

        // Group 5: 選択補助
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectSameFormat;

        // Group 6: ハンドル調整（新規追加）
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAdjustmentHandles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAngleHandles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetAdjustments;
    }

    partial class ThisRibbonCollection
    {
        //internal CustomRibbon CustomRibbon
        //{
        //    get { return this.GetRibbon<CustomRibbon>(); }
        //}
    }
}