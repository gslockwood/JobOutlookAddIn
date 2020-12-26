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
            this.version.Label = "version" + ThisAddIn.version;
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
            this.buttonForceEmail = this.Factory.CreateRibbonButton();
            this.version = this.Factory.CreateRibbonLabel();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.buttonConditionalEmail = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.buttonReset = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.buttonEditRoles = this.Factory.CreateRibbonButton();
            this.buttonEditCities = this.Factory.CreateRibbonButton();
            this.buttonEditOutgoingMessage = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.version);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.buttonConditionalEmail);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.buttonReset);
            this.group1.Items.Add(this.separator3);
            this.group1.Items.Add(this.buttonEditRoles);
            this.group1.Items.Add(this.buttonEditCities);
            this.group1.Items.Add(this.buttonEditOutgoingMessage);
            this.group1.Label = "Jobs Email Group";
            this.group1.Name = "group1";
            // 
            // buttonForceEmail
            // 
            this.buttonForceEmail.Label = "Force Send";
            this.buttonForceEmail.Name = "buttonForceEmail";
            this.buttonForceEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonForceEmail_Click);
            // 
            // version
            // 
            this.version.Label = "version";
            this.version.Name = "version";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // buttonConditionalEmail
            // 
            this.buttonConditionalEmail.Label = "Conditional Send";
            this.buttonConditionalEmail.Name = "buttonConditionalEmail";
            this.buttonConditionalEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonConditionalEmail_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // buttonReset
            // 
            this.buttonReset.Label = "Reset";
            this.buttonReset.Name = "buttonReset";
            this.buttonReset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonReset_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // buttonEditRoles
            // 
            this.buttonEditRoles.Label = "Edit Roles";
            this.buttonEditRoles.Name = "buttonEditRoles";
            this.buttonEditRoles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEditRoles_Click);
            // 
            // buttonEditCities
            // 
            this.buttonEditCities.Label = "Edit Cities";
            this.buttonEditCities.Name = "buttonEditCities";
            this.buttonEditCities.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEditCities_Click);
            // 
            // buttonEditOutgoingMessage
            // 
            this.buttonEditOutgoingMessage.Label = "Edit OutgoingMessage";
            this.buttonEditOutgoingMessage.Name = "buttonEditOutgoingMessage";
            this.buttonEditOutgoingMessage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEditOutgoingMessage_Click);
            // 
            // Job1Ribbon
            // 
            this.Name = "Job1Ribbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mai" +
    "l.Read";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.JobEmails);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel version;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEditRoles;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonConditionalEmail;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEditCities;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEditOutgoingMessage;
	}

	partial class ThisRibbonCollection
	{
		internal Job1Ribbon Ribbon1
		{
			get { return this.GetRibbon<Job1Ribbon>(); }
		}
	}
}
