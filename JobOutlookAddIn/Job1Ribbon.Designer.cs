namespace JobOutlookAddIn
{
	partial class Job1Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public Job1Ribbon()
			: base( Globals.Factory.GetRibbonFactory() )
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose( bool disposing )
		{
			if( disposing && ( components != null ) )
			{
				components.Dispose();
			}
			base.Dispose( disposing );
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tab1 = this.Factory.CreateRibbonTab();
			this.JobEmails = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.buttonReset = this.Factory.CreateRibbonButton();
			this.buttonConditionalEmail = this.Factory.CreateRibbonButton();
			this.buttonForceEmail = this.Factory.CreateRibbonButton();
			this.separator1 = this.Factory.CreateRibbonSeparator();
			this.separator2 = this.Factory.CreateRibbonSeparator();
			this.tab1.SuspendLayout();
			this.JobEmails.SuspendLayout();
			this.group1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Label = "TabAddIns";
			this.tab1.Name = "tab1";
			// 
			// JobEmails
			// 
			this.JobEmails.Groups.Add(this.group1);
			this.JobEmails.Label = "Job Emails";
			this.JobEmails.Name = "JobEmails";
			// 
			// group1
			// 
			this.group1.Items.Add(this.buttonForceEmail);
			this.group1.Items.Add(this.separator2);
			this.group1.Items.Add(this.buttonConditionalEmail);
			this.group1.Items.Add(this.separator1);
			this.group1.Items.Add(this.buttonReset);
			this.group1.Label = "Jobs Email Group";
			this.group1.Name = "group1";
			// 
			// buttonReset
			// 
			this.buttonReset.Label = "Reset";
			this.buttonReset.Name = "buttonReset";
			this.buttonReset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonReset_Click);
			// 
			// buttonConditionalEmail
			// 
			this.buttonConditionalEmail.Label = "Conditional Email";
			this.buttonConditionalEmail.Name = "buttonConditionalEmail";
			this.buttonConditionalEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonConditionalEmail_Click);
			// 
			// buttonForceEmail
			// 
			this.buttonForceEmail.Label = "Force Email";
			this.buttonForceEmail.Name = "buttonForceEmail";
			this.buttonForceEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonForceEmail_Click);
			// 
			// separator1
			// 
			this.separator1.Name = "separator1";
			// 
			// separator2
			// 
			this.separator2.Name = "separator2";
			// 
			// Job1Ribbon
			// 
			this.Name = "Job1Ribbon";
			this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mai" +
    "l.Read";
			this.Tabs.Add(this.tab1);
			this.Tabs.Add(this.JobEmails);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.JobEmails.ResumeLayout(false);
			this.JobEmails.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		private Microsoft.Office.Tools.Ribbon.RibbonTab JobEmails;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonReset;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonForceEmail;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonConditionalEmail;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
	}

	partial class ThisRibbonCollection
	{
		internal Job1Ribbon Ribbon1
		{
			get { return this.GetRibbon<Job1Ribbon>(); }
		}
	}
}
