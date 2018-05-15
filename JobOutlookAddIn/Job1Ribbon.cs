using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace JobOutlookAddIn
{
	public partial class Job1Ribbon
	{
		private void Ribbon1_Load( object sender, RibbonUIEventArgs e )
		{

		}



		private void buttonReset_Click( object sender, RibbonControlEventArgs e )
		{
			Globals.ThisAddIn.Reset();
		}

		private void buttonForceEmail_Click( object sender, RibbonControlEventArgs e )
		{
			Globals.ThisAddIn.ForceEmail();

		}

		private void buttonConditionalEmail_Click( object sender, RibbonControlEventArgs e )
		{
			Globals.ThisAddIn.ConditionalEmail();

		}
	}
}
