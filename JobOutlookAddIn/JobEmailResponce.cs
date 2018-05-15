using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

using Utilities;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace JobOutlookAddIn
{
	internal class JobEmailResponce : IJobEmailResponce
	{
		private readonly string attachmentFileName;
		private readonly string outGoingMessage;
		private readonly IList<Entity> roles;


		public JobEmailResponce( string fileName, string attachmentFileName, IList<Entity> roles )
		{
			outGoingMessage = FileUtilities.ReadFile( fileName );
			this.roles = roles;
			this.attachmentFileName = attachmentFileName;			
			//
		}

		public void ImmediateReply( object item, SendCondition condition )
		{
			Microsoft.Office.Interop.Outlook.MailItem mailItem = item as Microsoft.Office.Interop.Outlook.MailItem;
			if( mailItem == null )
				throw new Exception( "ImmediateReply: mailItem is undefined." );


			string body = CreateResponseBody( mailItem.Subject.ToLower(), mailItem.Sender.Name, condition );

			if( body == null )
				return;

			Microsoft.Office.Interop.Outlook.MailItem reply = mailItem.ReplyAll();
			reply.HTMLBody = AdjustText( reply.HTMLBody );

			reply.HTMLBody = body + reply.HTMLBody;

			reply.Attachments.Add( this.attachmentFileName );

			if( condition == SendCondition.Conditional )
				reply.Display();
			else
				reply.Send();

			//mailItem.HTMLBody = string.Format( "<font size='1' color='red'><div id='replysent' > Reply sent on: {0}</div><font><br>", System.DateTime.Now.ToShortDateString() ) + mailItem.HTMLBody;
			//font-size:14px; font-family:Times New Roman;
			mailItem.HTMLBody = string.Format( "<div id='replysent' style='color:darkblue; font-size:8px; font-family:Times New Roman;' > Reply sent on: {0}</div><br>", System.DateTime.Now.ToShortDateString() ) + mailItem.HTMLBody;


#if true5555
			string subject = mailItem.Subject.ToLower();
			string senderName = mailItem.Sender.Name;
			string senderAddress = mailItem.Sender.Address;


			string[] array = senderName.Split( ' ' );
			if( array.Length > 1 )
				senderName = array[0];
			else
			{
				array = senderName.Split( '.' );
				if( array.Length > 1 )
					senderName = array[0];
			}

			IList<Entity> relevantRoles = new List<Entity>();
			IList<Entity> irrelevantRoles = new List<Entity>();
			IList<Entity> veryIrrelevantRoles = new List<Entity>();

			foreach( Entity entity in roles )
			{
				if( isMatched( subject, entity.Item.ToLower() ) )
				//if (subject.Contains( entity.Item.ToLower() ) )
				{
					if( entity.Attrib == 0 )
						veryIrrelevantRoles.Add( new Entity( entity.Item, entity.Attrib ) );
					else if( entity.Attrib == 1 )
						irrelevantRoles.Add( new Entity( entity.Item, entity.Attrib ) );
					else if( entity.Attrib == 2 )
						relevantRoles.Add( new Entity( entity.Item, entity.Attrib ) );
					else
						throw new Exception( "data error: " + entity.ToString() );

				}


			}

			if( relevantRoles.Count > 0 && condition == SendCondition.Conditional )
			{
				//MarkAsProcessed
				System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show( "RelevantRole (" + relevantRoles[0].Item + ") found." );
				if( result == System.Windows.Forms.DialogResult.No )
					return;
				//
			}

			bool proceed = false;

			if( condition == SendCondition.Conditional )
			{
				if( ( irrelevantRoles.Count > 0 ) || ( veryIrrelevantRoles.Count > 0 ) )
					proceed = true;

			}
			else
				proceed = true;


			if( proceed )
			{
				string body = "<div style='font-size:14px; font-family:Times New Roman;'>" + this.outGoingMessage + "</div>";

				System.Globalization.TextInfo myTI = new System.Globalization.CultureInfo( "en-US", false ).TextInfo;
				body = body.Replace( "<person>", myTI.ToTitleCase( senderName ) );


				//if( ( irrelevantRoles.Count > 0 ) || ( veryIrrelevantRoles.Count > 0 ) )
				{
					if( ( veryIrrelevantRoles.Count > 0 ) )
						body = body.Replace( "<veryirrelevant>", "But really...." + veryIrrelevantRoles[0].Item + "??" );
					else
						body = body.Replace( "<veryirrelevant>", "" );

					if( irrelevantRoles.Count > 0 )
						body = body.Replace( "<role>", "(" + irrelevantRoles[0].Item + "s)" );
					else if( veryIrrelevantRoles.Count > 0 )
						body = body.Replace( "<role>", "(" + veryIrrelevantRoles[0].Item + "s)" );
					else
						body = body.Replace( "<role>", "" );

					//body = string.Format( "<font size='1' color='red'><div id='gslprocesed' > Processed on: {0} Found {1}</div><font>", System.DateTime.Now.ToShortDateString() ) + body;
					//body = string.Format( "<font size='1' color='red'><div id='replysent' > Reply sent on: {0}</div><font><br>", System.DateTime.Now.ToShortDateString() ) + body;

					Microsoft.Office.Interop.Outlook.MailItem reply = mailItem.ReplyAll();
					reply.HTMLBody = body + reply.HTMLBody;

					reply.Attachments.Add( this.attachmentFileName );

					if( condition == SendCondition.Conditional )
						reply.Display();
					else
						reply.Send();

					mailItem.HTMLBody = string.Format( "<font size='1' color='red'><div id='replysent' > Reply sent on: {0}</div><font><br>", System.DateTime.Now.ToShortDateString() ) + mailItem.HTMLBody;
					//
				}
				//
			} 
#endif

		}

		private string AdjustText( string hTMLBody )
		{
			zz:
			HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
			if( hTMLBody.Contains( "Regards," ))
			{
				doc.LoadHtml( hTMLBody );
				if( doc.DocumentNode.SelectNodes( "//a[@name='_MailAutoSig']" ) != null )
				{
					foreach( HtmlAgilityPack.HtmlNode node in doc.DocumentNode.SelectNodes( "//a[@name='_MailAutoSig']" ) )
						node.Remove();

				}
				if( doc.DocumentNode.SelectNodes( "//span[@style='mso-bookmark:_MailAutoSig']" ) != null )
				{
					foreach( HtmlAgilityPack.HtmlNode node in doc.DocumentNode.SelectNodes( "//span[@style='mso-bookmark:_MailAutoSig']" ) )
						node.Remove();

				}

				/*
				HtmlAgilityPack.HtmlNode node = doc.DocumentNode.SelectSingleNode( "//a[@name='_MailAutoSig']" );
				if( node != null )
					node.Remove();

				node = doc.DocumentNode.SelectSingleNode( "//span[@style='mso-bookmark:_MailAutoSig']" );
				if( node != null )
					node.Remove();
					*/
				//hTMLBody = hTMLBody.Replace( "Regards,", "" );
				//hTMLBody = hTMLBody.Replace( "george", "" );

			}
			//goto zz;

			return doc.DocumentNode.InnerHtml;
			//
		}

		private string CreateResponseBody( string subject, string senderName, SendCondition condition )
		{
			//string senderName = sender.Name;
			//string senderAddress = sender.Address;

			string[] array = senderName.Split( ' ' );
			if( array.Length > 1 )
				senderName = array[0];
			else
			{
				array = senderName.Split( '.' );
				if( array.Length > 1 )
					senderName = array[0];
			}

			IList<Entity> relevantRoles = new List<Entity>();
			IList<Entity> irrelevantRoles = new List<Entity>();
			IList<Entity> veryIrrelevantRoles = new List<Entity>();

			foreach( Entity entity in roles )
			{
				if( isMatched( subject, entity.Item.ToLower() ) )
				//if (subject.Contains( entity.Item.ToLower() ) )
				{
					if( entity.Attrib == 0 )
						veryIrrelevantRoles.Add( new Entity( entity.Item, entity.Attrib ) );
					else if( entity.Attrib == 1 )
						irrelevantRoles.Add( new Entity( entity.Item, entity.Attrib ) );
					else if( entity.Attrib == 2 )
						relevantRoles.Add( new Entity( entity.Item, entity.Attrib ) );
					else
						throw new Exception( "data error: " + entity.ToString() );
					//
				}

			}

			bool proceed = false;

			if( relevantRoles.Count > 0 && condition == SendCondition.Conditional )
			{
				DialogResult result = MessageBox.Show( "RelevantRole (" + relevantRoles[0].Item + ") found." + Environment.NewLine + Environment.NewLine + "Proceed?", "Proceed?", MessageBoxButtons.YesNo );
				if( result == System.Windows.Forms.DialogResult.No )
					return null;

				proceed = true;
				//
			}
			else
			{
				if( condition == SendCondition.Conditional )
				{
					if( ( irrelevantRoles.Count > 0 ) || ( veryIrrelevantRoles.Count > 0 ) )
						proceed = true;
					else
					{
						DialogResult result = System.Windows.Forms.MessageBox.Show( "Did not find an IrrelevantRole." + Environment.NewLine + Environment.NewLine + "Proceed?", "Proceed?", MessageBoxButtons.YesNo );
						if( result == System.Windows.Forms.DialogResult.No )
							return null;
						//
					}
					//
				}
				else
					proceed = true;
				//
			}

			if( proceed )
			{
				string body = "<div style='font-size:14px; font-family:Times New Roman;'>" + this.outGoingMessage + "</div>";

				System.Globalization.TextInfo myTI = new System.Globalization.CultureInfo( "en-US", false ).TextInfo;
				body = body.Replace( "<person>", myTI.ToTitleCase( senderName ) );

				if( ( veryIrrelevantRoles.Count > 0 ) )
					body = body.Replace( "<veryirrelevant>", "But really...." + veryIrrelevantRoles[0].Item + "??" );
				else
					body = body.Replace( "<veryirrelevant>", "" );

				if( irrelevantRoles.Count > 0 )
					body = body.Replace( "<role>", "(" + irrelevantRoles[0].Item + "s)" );
				else if( veryIrrelevantRoles.Count > 0 )
					body = body.Replace( "<role>", "(" + veryIrrelevantRoles[0].Item + "s)" );
				else
					body = body.Replace( "<role>", "" );


				return body;
				//
			}

			return null;
			//
		}

		private bool isMatched( string subject, string searchString )
		{
aaa:
//(engineer)[ ,.;0-9]
			string pattern = "(" + searchString + ")\\b";//[ ,.;0-9]";

			var regex = new Regex( pattern );
			Match match = regex.Match( subject );
			//bool restul = match.Success;
			//goto aaa;
			return match.Success;
			//
		}
	}
}
/*
				try
				{
				}
				catch (Exception ex)
				{
					//throw ex;
				}
				*/