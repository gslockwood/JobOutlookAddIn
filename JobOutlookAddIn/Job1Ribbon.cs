using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace JobOutlookAddIn
{
	public partial class Job1Ribbon
	{
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

		private void buttonEditRoles_Click( object sender, RibbonControlEventArgs e )
		{
			Globals.ThisAddIn.EditFile( ThisAddIn.FileCategory.Role );
		}
		private void buttonEditCities_Click( object sender, RibbonControlEventArgs e )
		{
			Globals.ThisAddIn.EditFile( ThisAddIn.FileCategory.City );

		}

		private void buttonEditOutgoingMessage_Click( object sender, RibbonControlEventArgs e )
		{
			Globals.ThisAddIn.EditFile( ThisAddIn.FileCategory.OutGoingMessage );

		}
	}
}
