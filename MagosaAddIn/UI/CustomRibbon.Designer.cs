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
            this.btnLayerAdjustment = this.Factory.CreateRibbonButton();
            this.btnAutoNumbering = this.Factory.CreateRibbonButton();
            this.btnThemeColorGenerator = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnAlignToLeft = this.Factory.CreateRibbonButton();
            this.btnAlignToRight = this.Factory.CreateRibbonButton();
            this.btnAlignToTop = this.Factory.CreateRibbonButton();
            this.btnAlignToBottom = this.Factory.CreateRibbonButton();
            this.btnAlignToHorizontalCenter = this.Factory.CreateRibbonButton();
            this.btnAlignToVerticalCenter = this.Factory.CreateRibbonButton();
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
            // 新規追加: 図形置き換えグループ
            this.group7 = this.Factory.CreateRibbonGroup();
            this.btnSaveShapes = this.Factory.CreateRibbonButton();
            this.btnReplaceShapes = this.Factory.CreateRibbonButton();
            this.lblSavedCount = this.Factory.CreateRibbonLabel();
            // 新規追加: サイズ調整グループ
            this.group8 = this.Factory.CreateRibbonGroup();
            this.menuSizeUnify = this.Factory.CreateRibbonMenu();
            this.btnResizeToReference = this.Factory.CreateRibbonButton();
            this.btnResizeToMaximum = this.Factory.CreateRibbonButton();
            this.btnResizeToMinimum = this.Factory.CreateRibbonButton();
            this.menuKeepRatio = this.Factory.CreateRibbonMenu();
            this.btnResizeToWidthKeepRatio = this.Factory.CreateRibbonButton();
            this.btnResizeToHeightKeepRatio = this.Factory.CreateRibbonButton();
            this.btnResizeDialog = this.Factory.CreateRibbonButton();
            // 新規追加: 配列複製グループ
            this.group9 = this.Factory.CreateRibbonGroup();
            this.btnCircularArray = this.Factory.CreateRibbonButton();
            this.btnGridArray = this.Factory.CreateRibbonButton();

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
            this.tab1.Groups.Add(this.group5); // ①選択補助
            this.tab1.Groups.Add(this.group1); // ②図形操作
            this.tab1.Groups.Add(this.group7); // ③図形置き換え
            this.tab1.Groups.Add(this.group6); // ④ハンドル調整
            this.tab1.Groups.Add(this.group8); // ⑤サイズ調整
            this.tab1.Groups.Add(this.group9); // ⑥配列複製
            this.tab1.Groups.Add(this.group2); // ⑦基準整列
            this.tab1.Groups.Add(this.group3); // ⑧隣接整列
            this.tab1.Groups.Add(this.group4); // ⑨拡張整列
            this.tab1.Label = "Magosa Tools";
            this.tab1.Name = "tab1";

            // 
            // group1
            // 
            this.group1.Items.Add(this.btnDivideShape);
            this.group1.Items.Add(this.btnLayerAdjustment);
            this.group1.Items.Add(this.btnAutoNumbering);
            this.group1.Items.Add(this.btnThemeColorGenerator);
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
            // btnLayerAdjustment
            // 
            this.btnLayerAdjustment.Label = "レイヤー調整";
            this.btnLayerAdjustment.Name = "btnLayerAdjustment";
            this.btnLayerAdjustment.OfficeImageId = "ObjectBringToFront";
            this.btnLayerAdjustment.ShowImage = true;
            this.btnLayerAdjustment.SuperTip = "選択した図形の重なり順を調整します。選択順または位置に基づいて前面・背面を設定できます。";
            this.btnLayerAdjustment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLayerAdjustment_Click);

            // 
            // btnAutoNumbering
            // 
            this.btnAutoNumbering.Label = "自動ナンバリング";
            this.btnAutoNumbering.Name = "btnAutoNumbering";
            this.btnAutoNumbering.OfficeImageId = "NumberStyleGallery";
            this.btnAutoNumbering.ShowImage = true;
            this.btnAutoNumbering.SuperTip = "選択した図形に自動で番号を付けます。算用数字、丸数字、アルファベット、ローマ数字などから選択できます。";
            this.btnAutoNumbering.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAutoNumbering_Click);

            // 
            // btnThemeColorGenerator
            // 
            this.btnThemeColorGenerator.Label = "カラー生成";
            this.btnThemeColorGenerator.Name = "btnThemeColorGenerator";
            this.btnThemeColorGenerator.OfficeImageId = "SmartArtChangeColorsGallery";
            this.btnThemeColorGenerator.ShowImage = true;
            this.btnThemeColorGenerator.SuperTip = "配色理論に基づいて17種類のカラーパレットを生成します。図形への適用やスライド枠外への配置が可能です。";
            this.btnThemeColorGenerator.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnThemeColorGenerator_Click);

            // 
            // group2
            // 
            this.group2.Items.Add(this.btnAlignToLeft);
            this.group2.Items.Add(this.btnAlignToRight);
            this.group2.Items.Add(this.btnAlignToTop);
            this.group2.Items.Add(this.btnAlignToBottom);
            this.group2.Items.Add(this.btnAlignToHorizontalCenter);
            this.group2.Items.Add(this.btnAlignToVerticalCenter);
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
            // btnAlignToHorizontalCenter
            // 
            this.btnAlignToHorizontalCenter.Label = "水平中央揃え";
            this.btnAlignToHorizontalCenter.Name = "btnAlignToHorizontalCenter";
            this.btnAlignToHorizontalCenter.OfficeImageId = "AlignDistributeHorizontally";
            this.btnAlignToHorizontalCenter.ShowImage = true;
            this.btnAlignToHorizontalCenter.SuperTip = "基準図形の水平中央に、その他の図形の水平中央を揃えます。";
            this.btnAlignToHorizontalCenter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToHorizontalCenter_Click);

            // 
            // btnAlignToVerticalCenter
            // 
            this.btnAlignToVerticalCenter.Label = "垂直中央揃え";
            this.btnAlignToVerticalCenter.Name = "btnAlignToVerticalCenter";
            this.btnAlignToVerticalCenter.OfficeImageId = "AlignDistributeVertically";
            this.btnAlignToVerticalCenter.ShowImage = true;
            this.btnAlignToVerticalCenter.SuperTip = "基準図形の垂直中央に、その他の図形の垂直中央を揃えます。";
            this.btnAlignToVerticalCenter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignToVerticalCenter_Click);

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
            this.btnAlignAndDistributeHorizontal.OfficeImageId = "HorizontalSpacingIncrease";
            this.btnAlignAndDistributeHorizontal.ShowImage = true;
            this.btnAlignAndDistributeHorizontal.SuperTip = "基準図形の水平中央に揃えて等間隔で配置します。2つの図形の場合は中央揃えのみ実行されます。";
            this.btnAlignAndDistributeHorizontal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignAndDistributeHorizontal_Click);

            // 
            // btnAlignAndDistributeVertical
            // 
            this.btnAlignAndDistributeVertical.Label = "垂直中央・等間隔";
            this.btnAlignAndDistributeVertical.Name = "btnAlignAndDistributeVertical";
            this.btnAlignAndDistributeVertical.OfficeImageId = "VerticalSpacingIncrease";
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
            // group7
            // 
            this.group7.Items.Add(this.btnSaveShapes);
            this.group7.Items.Add(this.btnReplaceShapes);
            this.group7.Items.Add(this.lblSavedCount);
            this.group7.Label = "図形置き換え";
            this.group7.Name = "group7";

            // 
            // btnSaveShapes
            // 
            this.btnSaveShapes.Label = "選択完了";
            this.btnSaveShapes.Name = "btnSaveShapes";
            this.btnSaveShapes.OfficeImageId = "AreaSelect";
            this.btnSaveShapes.ShowImage = true;
            this.btnSaveShapes.SuperTip = "置き換える複数の図形を選択して記憶します。記憶後、テンプレート図形を選択して置き換え実行してください。";
            this.btnSaveShapes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveShapes_Click);

            // 
            // btnReplaceShapes
            // 
            this.btnReplaceShapes.Label = "置き換え実行";
            this.btnReplaceShapes.Name = "btnReplaceShapes";
            this.btnReplaceShapes.OfficeImageId = "ReplaceShape";
            this.btnReplaceShapes.ShowImage = true;
            this.btnReplaceShapes.SuperTip = "記憶した図形をテンプレート図形で一括置き換えます。テンプレート図形を1つ選択してからクリックしてください。";
            this.btnReplaceShapes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplaceShapes_Click);

            // 
            // lblSavedCount
            // 
            this.lblSavedCount.Label = "記憶: 0個";
            this.lblSavedCount.Name = "lblSavedCount";

            // 
            // group8
            // 
            this.group8.Items.Add(this.btnResizeDialog);
            this.group8.Items.Add(this.menuSizeUnify);
            this.group8.Items.Add(this.menuKeepRatio);
            this.group8.Label = "サイズ調整";
            this.group8.Name = "group8";

            // 
            // menuSizeUnify
            // 
            this.menuSizeUnify.Items.Add(this.btnResizeToReference);
            this.menuSizeUnify.Items.Add(this.btnResizeToMaximum);
            this.menuSizeUnify.Items.Add(this.btnResizeToMinimum);
            this.menuSizeUnify.Label = "サイズ統一";
            this.menuSizeUnify.Name = "menuSizeUnify";
            this.menuSizeUnify.OfficeImageId = "SizeToControlHeightAndWidth";
            this.menuSizeUnify.ShowImage = true;

            // 
            // menuKeepRatio
            // 
            this.menuKeepRatio.Items.Add(this.btnResizeToWidthKeepRatio);
            this.menuKeepRatio.Items.Add(this.btnResizeToHeightKeepRatio);
            this.menuKeepRatio.Label = "比率保持";
            this.menuKeepRatio.Name = "menuKeepRatio";
            this.menuKeepRatio.OfficeImageId = "DiagramScale";
            this.menuKeepRatio.ShowImage = true;

            // 
            // btnResizeToReference
            // 
            this.btnResizeToReference.Label = "基準サイズに合わせる";
            this.btnResizeToReference.Name = "btnResizeToReference";
            this.btnResizeToReference.OfficeImageId = "AutomaticResize";
            this.btnResizeToReference.ShowImage = true;
            this.btnResizeToReference.SuperTip = "選択した複数の図形を、基準図形と同じサイズに調整します。";
            this.btnResizeToReference.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResizeToReference_Click);

            // 
            // btnResizeToWidthKeepRatio
            // 
            this.btnResizeToWidthKeepRatio.Label = "幅を統一";
            this.btnResizeToWidthKeepRatio.Name = "btnResizeToWidthKeepRatio";
            this.btnResizeToWidthKeepRatio.OfficeImageId = "SizeToControlWidth";
            this.btnResizeToWidthKeepRatio.ShowImage = true;
            this.btnResizeToWidthKeepRatio.SuperTip = "選択した図形の幅を統一し、アスペクト比を保ったまま高さを調整します。";
            this.btnResizeToWidthKeepRatio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResizeToWidthKeepRatio_Click);

            // 
            // btnResizeToHeightKeepRatio
            // 
            this.btnResizeToHeightKeepRatio.Label = "高さを統一";
            this.btnResizeToHeightKeepRatio.Name = "btnResizeToHeightKeepRatio";
            this.btnResizeToHeightKeepRatio.OfficeImageId = "SizeToControlHeight";
            this.btnResizeToHeightKeepRatio.ShowImage = true;
            this.btnResizeToHeightKeepRatio.SuperTip = "選択した図形の高さを統一し、アスペクト比を保ったまま幅を調整します。";
            this.btnResizeToHeightKeepRatio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResizeToHeightKeepRatio_Click);

            // 
            // btnResizeToMaximum
            // 
            this.btnResizeToMaximum.Label = "最大サイズに合わせる";
            this.btnResizeToMaximum.Name = "btnResizeToMaximum";
            this.btnResizeToMaximum.OfficeImageId = "SizeToControlHeightAndWidth";
            this.btnResizeToMaximum.ShowImage = true;
            this.btnResizeToMaximum.SuperTip = "選択した図形を最大サイズの図形に合わせてサイズを統一します。";
            this.btnResizeToMaximum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResizeToMaximum_Click);

            // 
            // btnResizeToMinimum
            // 
            this.btnResizeToMinimum.Label = "最小サイズに合わせる";
            this.btnResizeToMinimum.Name = "btnResizeToMinimum";
            this.btnResizeToMinimum.OfficeImageId = "SizeToControlHeightAndWidth";
            this.btnResizeToMinimum.ShowImage = true;
            this.btnResizeToMinimum.SuperTip = "選択した図形を最小サイズの図形に合わせてサイズを統一します。";
            this.btnResizeToMinimum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResizeToMinimum_Click);

            // 
            // btnResizeDialog
            // 
            this.btnResizeDialog.Label = "サイズ指定";
            this.btnResizeDialog.Name = "btnResizeDialog";
            this.btnResizeDialog.OfficeImageId = "SizeToFit";
            this.btnResizeDialog.ShowImage = true;
            this.btnResizeDialog.SuperTip = "パーセント指定での拡大縮小、または固定サイズ（mm/cm/pt）への調整をダイアログで設定します。";
            this.btnResizeDialog.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResizeDialog_Click);

            // 
            // group9
            // 
            this.group9.Items.Add(this.btnCircularArray);
            this.group9.Items.Add(this.btnGridArray);
            this.group9.Label = "配列複製";
            this.group9.Name = "group9";

            // 
            // btnCircularArray
            //
            this.btnCircularArray.Label = "円形配列";
            this.btnCircularArray.Name = "btnCircularArray";
            this.btnCircularArray.OfficeImageId = "DiagramRadialInsertClassic";
            this.btnCircularArray.ShowImage = true;
            this.btnCircularArray.SuperTip = "選択した図形を指定した中心・半径・個数で円形に配列します。";
            this.btnCircularArray.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCircularArray_Click);

            // 
            // btnGridArray
            // 
            this.btnGridArray.Label = "グリッド配列";
            this.btnGridArray.Name = "btnGridArray";
            this.btnGridArray.OfficeImageId = "SmartAlign";
            this.btnGridArray.ShowImage = true;
            this.btnGridArray.SuperTip = "選択した図形を指定した行×列でグリッド状に配列します。";
            this.btnGridArray.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGridArray_Click);

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

        // Group 1: 図形操作
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDivideShape;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLayerAdjustment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAutoNumbering;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemeColorGenerator;

        // Group 2: 基準整列
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToHorizontalCenter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignToVerticalCenter;

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

        // Group 7: 図形置き換え（新規追加）
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveShapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplaceShapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSavedCount;

        // Group 8: サイズ調整（新規追加）
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group8;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuSizeUnify;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResizeToReference;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResizeToMaximum;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResizeToMinimum;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuKeepRatio;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResizeToWidthKeepRatio;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResizeToHeightKeepRatio;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResizeDialog;

        // Group 9: 配列複製（新規追加）
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCircularArray;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGridArray;
    }

    partial class ThisRibbonCollection
    {
        //internal CustomRibbon CustomRibbon
        //{
        //    get { return this.GetRibbon<CustomRibbon>(); }
        //}
    }
}