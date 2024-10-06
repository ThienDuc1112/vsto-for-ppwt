namespace PowerPointAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl13 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl14 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnInsertNewSlide = this.Factory.CreateRibbonButton();
            this.btnCopySlide = this.Factory.CreateRibbonButton();
            this.btnAutomate = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnAddNote = this.Factory.CreateRibbonButton();
            this.btnExport = this.Factory.CreateRibbonButton();
            this.btnImportImages = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnApplyEffect = this.Factory.CreateRibbonButton();
            this.btnSaveSlide = this.Factory.CreateRibbonButton();
            this.btnTableTitle = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnSearch = this.Factory.CreateRibbonButton();
            this.btnCallApi = this.Factory.CreateRibbonButton();
            this.ddCity = this.Factory.CreateRibbonDropDown();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.btnFormatImage = this.Factory.CreateRibbonButton();
            this.btnTriggleVisibility = this.Factory.CreateRibbonToggleButton();
            this.btnAddText = this.Factory.CreateRibbonButton();
            this.btnRemoveShape = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.btnChangeFS = this.Factory.CreateRibbonButton();
            this.btnSplit = this.Factory.CreateRibbonButton();
            this.btnRemoveSlide = this.Factory.CreateRibbonButton();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.ddBackgroundSelecting = this.Factory.CreateRibbonDropDown();
            this.btnApplyBg = this.Factory.CreateRibbonButton();
            this.btnReverse = this.Factory.CreateRibbonButton();
            this.DocumentOperation = this.Factory.CreateRibbonTab();
            this.group8 = this.Factory.CreateRibbonGroup();
            this.btnSplitPP = this.Factory.CreateRibbonButton();
            this.btnCombinePp = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.group6.SuspendLayout();
            this.group7.SuspendLayout();
            this.DocumentOperation.SuspendLayout();
            this.group8.SuspendLayout();
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
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group7);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnInsertNewSlide);
            this.group1.Items.Add(this.btnCopySlide);
            this.group1.Items.Add(this.btnAutomate);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnInsertNewSlide
            // 
            this.btnInsertNewSlide.Label = "Insert New Slide";
            this.btnInsertNewSlide.Name = "btnInsertNewSlide";
            this.btnInsertNewSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertNewSlide_Click);
            // 
            // btnCopySlide
            // 
            this.btnCopySlide.Label = "Duplicate";
            this.btnCopySlide.Name = "btnCopySlide";
            this.btnCopySlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopySlide_Click);
            // 
            // btnAutomate
            // 
            this.btnAutomate.Label = "Automate";
            this.btnAutomate.Name = "btnAutomate";
            this.btnAutomate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAutomate_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnAddNote);
            this.group2.Items.Add(this.btnExport);
            this.group2.Items.Add(this.btnImportImages);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // btnAddNote
            // 
            this.btnAddNote.Label = "Add Note";
            this.btnAddNote.Name = "btnAddNote";
            this.btnAddNote.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddNote_Click);
            // 
            // btnExport
            // 
            this.btnExport.Label = "Save as PDF";
            this.btnExport.Name = "btnExport";
            this.btnExport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExport_Click);
            // 
            // btnImportImages
            // 
            this.btnImportImages.Label = "Import Image";
            this.btnImportImages.Name = "btnImportImages";
            this.btnImportImages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportImages_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnApplyEffect);
            this.group3.Items.Add(this.btnSaveSlide);
            this.group3.Items.Add(this.btnTableTitle);
            this.group3.Label = "group3";
            this.group3.Name = "group3";
            // 
            // btnApplyEffect
            // 
            this.btnApplyEffect.Label = "Apply Effect";
            this.btnApplyEffect.Name = "btnApplyEffect";
            this.btnApplyEffect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnApplyEffect_Click);
            // 
            // btnSaveSlide
            // 
            this.btnSaveSlide.Label = "Save Current Slide";
            this.btnSaveSlide.Name = "btnSaveSlide";
            this.btnSaveSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveSlide_Click);
            // 
            // btnTableTitle
            // 
            this.btnTableTitle.Label = "Create TOC";
            this.btnTableTitle.Name = "btnTableTitle";
            this.btnTableTitle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTableTitle_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnSearch);
            this.group4.Items.Add(this.btnCallApi);
            this.group4.Items.Add(this.ddCity);
            this.group4.Label = "group4";
            this.group4.Name = "group4";
            // 
            // btnSearch
            // 
            this.btnSearch.Label = "Search";
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSearch_Click);
            // 
            // btnCallApi
            // 
            this.btnCallApi.Label = "Get Weather";
            this.btnCallApi.Name = "btnCallApi";
            this.btnCallApi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCallApi_Click);
            // 
            // ddCity
            // 
            ribbonDropDownItemImpl1.Label = "Berlin";
            ribbonDropDownItemImpl2.Label = "Paris";
            ribbonDropDownItemImpl3.Label = "Tokyo";
            ribbonDropDownItemImpl4.Label = "New York";
            ribbonDropDownItemImpl5.Label = "Sydney";
            ribbonDropDownItemImpl6.Label = "Mumbai";
            ribbonDropDownItemImpl7.Label = "Ha Noi";
            this.ddCity.Items.Add(ribbonDropDownItemImpl1);
            this.ddCity.Items.Add(ribbonDropDownItemImpl2);
            this.ddCity.Items.Add(ribbonDropDownItemImpl3);
            this.ddCity.Items.Add(ribbonDropDownItemImpl4);
            this.ddCity.Items.Add(ribbonDropDownItemImpl5);
            this.ddCity.Items.Add(ribbonDropDownItemImpl6);
            this.ddCity.Items.Add(ribbonDropDownItemImpl7);
            this.ddCity.Label = "Select a city";
            this.ddCity.Name = "ddCity";
            // 
            // group5
            // 
            this.group5.Items.Add(this.btnFormatImage);
            this.group5.Items.Add(this.btnTriggleVisibility);
            this.group5.Items.Add(this.btnAddText);
            this.group5.Items.Add(this.btnRemoveShape);
            this.group5.Label = "group5";
            this.group5.Name = "group5";
            // 
            // btnFormatImage
            // 
            this.btnFormatImage.Label = "Format Image";
            this.btnFormatImage.Name = "btnFormatImage";
            this.btnFormatImage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatImage_Click);
            // 
            // btnTriggleVisibility
            // 
            this.btnTriggleVisibility.Label = "Triggle Visibility";
            this.btnTriggleVisibility.Name = "btnTriggleVisibility";
            this.btnTriggleVisibility.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTriggleVisibility_Click);
            // 
            // btnAddText
            // 
            this.btnAddText.Label = "Add Text";
            this.btnAddText.Name = "btnAddText";
            this.btnAddText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddText_Click);
            // 
            // btnRemoveShape
            // 
            this.btnRemoveShape.Label = "";
            this.btnRemoveShape.Name = "btnRemoveShape";
            // 
            // group6
            // 
            this.group6.Items.Add(this.btnChangeFS);
            this.group6.Items.Add(this.btnSplit);
            this.group6.Items.Add(this.btnRemoveSlide);
            this.group6.Label = "group6";
            this.group6.Name = "group6";
            // 
            // btnChangeFS
            // 
            this.btnChangeFS.Label = "ChangeFontStyle";
            this.btnChangeFS.Name = "btnChangeFS";
            this.btnChangeFS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeFS_Click);
            // 
            // btnSplit
            // 
            this.btnSplit.Label = "Split Into 2 Column";
            this.btnSplit.Name = "btnSplit";
            this.btnSplit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSplit_Click);
            // 
            // btnRemoveSlide
            // 
            this.btnRemoveSlide.Label = "Remove Slide";
            this.btnRemoveSlide.Name = "btnRemoveSlide";
            this.btnRemoveSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveSlide_Click);
            // 
            // group7
            // 
            this.group7.Items.Add(this.ddBackgroundSelecting);
            this.group7.Items.Add(this.btnApplyBg);
            this.group7.Items.Add(this.btnReverse);
            this.group7.Label = "group7";
            this.group7.Name = "group7";
            // 
            // ddBackgroundSelecting
            // 
            ribbonDropDownItemImpl8.Label = "Horizontal Gradient";
            ribbonDropDownItemImpl9.Label = "Vertical Gradient";
            ribbonDropDownItemImpl10.Label = "Diagonal Gradient";
            ribbonDropDownItemImpl11.Label = "Rectangular Gradient";
            ribbonDropDownItemImpl12.Label = "Path Gradient";
            ribbonDropDownItemImpl13.Label = "Mixed Gradient";
            ribbonDropDownItemImpl14.Label = "Center Gradient";
            this.ddBackgroundSelecting.Items.Add(ribbonDropDownItemImpl8);
            this.ddBackgroundSelecting.Items.Add(ribbonDropDownItemImpl9);
            this.ddBackgroundSelecting.Items.Add(ribbonDropDownItemImpl10);
            this.ddBackgroundSelecting.Items.Add(ribbonDropDownItemImpl11);
            this.ddBackgroundSelecting.Items.Add(ribbonDropDownItemImpl12);
            this.ddBackgroundSelecting.Items.Add(ribbonDropDownItemImpl13);
            this.ddBackgroundSelecting.Items.Add(ribbonDropDownItemImpl14);
            this.ddBackgroundSelecting.Label = "Select Background";
            this.ddBackgroundSelecting.Name = "ddBackgroundSelecting";
            // 
            // btnApplyBg
            // 
            this.btnApplyBg.Label = "Apply Background";
            this.btnApplyBg.Name = "btnApplyBg";
            this.btnApplyBg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnApplyBg_Click);
            // 
            // btnReverse
            // 
            this.btnReverse.Label = "Reverse Slides";
            this.btnReverse.Name = "btnReverse";
            this.btnReverse.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReverse_Click);
            // 
            // DocumentOperation
            // 
            this.DocumentOperation.Groups.Add(this.group8);
            this.DocumentOperation.Label = "Document Operation";
            this.DocumentOperation.Name = "DocumentOperation";
            // 
            // group8
            // 
            this.group8.Items.Add(this.btnSplitPP);
            this.group8.Items.Add(this.btnCombinePp);
            this.group8.Label = "group8";
            this.group8.Name = "group8";
            // 
            // btnSplitPP
            // 
            this.btnSplitPP.Label = "Split Powerpoint";
            this.btnSplitPP.Name = "btnSplitPP";
            this.btnSplitPP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSplitPP_Click);
            // 
            // btnCombinePp
            // 
            this.btnCombinePp.Label = "Combine 2 files";
            this.btnCombinePp.Name = "btnCombinePp";
            this.btnCombinePp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCombinePp_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.DocumentOperation);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
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
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.DocumentOperation.ResumeLayout(false);
            this.DocumentOperation.PerformLayout();
            this.group8.ResumeLayout(false);
            this.group8.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertNewSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopySlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAutomate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddNote;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportImages;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApplyEffect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTableTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCallApi;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddCity;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveShape;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTriggleVisibility;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddText;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeFS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSplit;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group7;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddBackgroundSelecting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApplyBg;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReverse;
        private Microsoft.Office.Tools.Ribbon.RibbonTab DocumentOperation;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSplitPP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCombinePp;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
