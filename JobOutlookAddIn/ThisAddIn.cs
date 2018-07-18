using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Utilities;
using HtmlAgilityPack;
using System.Runtime.Serialization;
using System.Net;

namespace JobOutlookAddIn
{
	public partial class ThisAddIn
	{
		Outlook.Inspectors inspectors;
		Outlook.Explorer currentExplorer = null;

		IList<Entity> roles;
		IList<Entity> cities;
		JobEmailResponce jobEmailResponce = null;

		//private IList<string> wrongTitlesList = null;


		private void ThisAddIn_Startup( object sender, System.EventArgs e )
		{
			inspectors = this.Application.Inspectors;
			//inspectors.NewInspector += Inspectors_NewInspector;

			currentExplorer = this.Application.ActiveExplorer();
			currentExplorer.SelectionChange += CurrentExplorer_SelectionChange;
			/*
			wrongTitlesList = new List<string>();
			wrongTitlesList.Add( "sales" );
			wrongTitlesList.Add( "hr " );
			wrongTitlesList.Add( "assistant" );
			wrongTitlesList.Add( "art" );
			wrongTitlesList.Add( "regional" );
			wrongTitlesList.Add( "marketing" );
			wrongTitlesList.Add( "culinary" );
			wrongTitlesList.Add( "investment" );
			wrongTitlesList.Add( "community" );
			wrongTitlesList.Add( "music" );
			wrongTitlesList.Add( "consulting" );
			wrongTitlesList.Add( "store" );
			wrongTitlesList.Add( "clinical" );
			wrongTitlesList.Add( "business" );
			wrongTitlesList.Add( "finance" );
			wrongTitlesList.Add( "director of development" );
			wrongTitlesList.Add( "operations" );
			wrongTitlesList.Add( "engagement" );
			wrongTitlesList.Add( "coordinator" );
			wrongTitlesList.Add( "sourcing" );
			wrongTitlesList.Add( "camp " );
			wrongTitlesList.Add( "administrative" );
			wrongTitlesList.Add( "government" );
			wrongTitlesList.Add( "controller" );
			wrongTitlesList.Add( "medical" );
			wrongTitlesList.Add( "program" );
			wrongTitlesList.Add( "spa " );
			wrongTitlesList.Add( "legal  " );
			wrongTitlesList.Add( "of development" );
			wrongTitlesList.Add( "training " );
			wrongTitlesList.Add( "regulatory " );
			wrongTitlesList.Add( "training " );
			*/
			ProcessTextfiles();
			//
		}

		internal void DoIt( object selObject, SendCondition sendCondition )
		{
			jobEmailResponce.ImmediateReply( selObject, sendCondition );

		}

		internal void ForceEmail()
		{
			try
			{
				if( this.Application.ActiveExplorer().Selection.Count > 0 )
					foreach( object selObject in this.Application.ActiveExplorer().Selection )
						if( selObject is Outlook.MailItem )
							jobEmailResponce.ImmediateReply( selObject, SendCondition.Force );
				//
			}
			catch( Exception ex )
			{
				MessageBox.Show( ex.Message );
			}
			//
		}

		internal void ConditionalEmail()
		{
			try
			{
				if( this.Application.ActiveExplorer().Selection.Count > 0 )
					foreach( object selObject in this.Application.ActiveExplorer().Selection )
						if( selObject is Outlook.MailItem )
							jobEmailResponce.ImmediateReply( selObject, SendCondition.Conditional );
				//
			}
			catch( Exception ex )
			{
				MessageBox.Show( ex.Message );
			}
			//
		}

		internal void Reset()
		{
			ProcessTextfiles();
		}


		private void ProcessTextfiles()
		{
			//string roles;
			//string cities;

			try
			{
				string userDir = Environment.GetFolderPath( Environment.SpecialFolder.UserProfile );
				//C:\Users\Georg\OneDrive\OutlookHelper
				roles = FileUtilities.ParseFile( userDir + @"\OneDrive\OutlookHelper\Roles.txt" );
				cities = FileUtilities.ParseFile( userDir + @"\OneDrive\OutlookHelper\Cities.txt" );
				jobEmailResponce = new JobEmailResponce( userDir + @"\OneDrive\OutlookHelper\outgoingmessage.txt", userDir + @"\OneDrive\Resumes\Current resumev.pdf", roles );

			}
			catch( Exception x )
			{
				MessageBox.Show( x.Message );
			}
			//
		}

		private void CurrentExplorer_SelectionChange()
		{
			Outlook.MAPIFolder selectedFolder = this.Application.ActiveExplorer().CurrentFolder;
			String expMessage = "Your current folder is " + selectedFolder.Name + ".\n";
			String itemMessage = "Item is unknown.";
			try
			{
				if( this.Application.ActiveExplorer().Selection.Count > 0 )
				{
					Object selObject = this.Application.ActiveExplorer().Selection[1];
					if( selObject is Outlook.MailItem )
					{
						Outlook.MailItem mailItem = ( selObject as Outlook.MailItem );

						if( mailItem.Sender.Address != "googlealerts-noreply@google.com" )
							ProcessEmail( mailItem );

						//itemMessage = "The item is an e-mail message." + " The subject is " + mailItem.Subject + ".";
						//mailItem.Display(false);
						//
					}
					else if( selObject is Outlook.ContactItem )
					{
						Outlook.ContactItem contactItem =
							( selObject as Outlook.ContactItem );
						itemMessage = "The item is a contact." +
							" The full name is " + contactItem.Subject + ".";
						contactItem.Display( false );
					}
					else if( selObject is Outlook.AppointmentItem )
					{
						Outlook.AppointmentItem apptItem =
							( selObject as Outlook.AppointmentItem );
						itemMessage = "The item is an appointment." +
							" The subject is " + apptItem.Subject + ".";
					}
					else if( selObject is Outlook.TaskItem )
					{
						Outlook.TaskItem taskItem =
							( selObject as Outlook.TaskItem );
						itemMessage = "The item is a task. The body is "
							+ taskItem.Body + ".";
					}
					else if( selObject is Outlook.MeetingItem )
					{
						Outlook.MeetingItem meetingItem =
							( selObject as Outlook.MeetingItem );
						itemMessage = "The item is a meeting item. " +
							 "The subject is " + meetingItem.Subject + ".";
					}
				}
				expMessage = expMessage + itemMessage;
			}
			catch( Exception ex )
			{
				expMessage = ex.Message;
			}

			//MessageBox.Show(expMessage);

		}

		private void ProcessEmail( Outlook.MailItem mailItem )
		{
			string htmlBody = mailItem.HTMLBody;

			try
			{
				var result = ProcessBody( htmlBody );
				if( result != null )
					mailItem.HTMLBody = result;


			}
			catch( NoActiveJobsException ex )
			{
				mailItem.Delete();
				//
			}
			catch( Exception ex )
			{
				//
			}
			//
			//
		}

		private string ProcessBody( string htmlBody )
		{
			//return null;
			//jim:
			//System.Drawing.Color c = System.Drawing.ColorTranslator.FromHtml("#F5F7F8");
			//String strHtmlColor = System.Drawing.ColorTranslator.ToHtml(c);
			////goto jim;

			///// if 
			if( string.IsNullOrEmpty( htmlBody ) )
				return null;

			HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
			doc.LoadHtml( htmlBody );

			if( doc.GetElementbyId( "gslprocesed" ) != null )
				return null;

			var div = from lnks in doc.DocumentNode.Descendants()
					  where lnks.Name == "span" && lnks.InnerText.Contains( "Fresh Jobs" )
					  select new
					  {
						  spanOfInterest = lnks
					  };

			HtmlNode spanFreshJobs = null;
			if( div.Count() > 0 )
			{
				spanFreshJobs = div.First().spanOfInterest;
				//System.Diagnostics.Debug.WriteLine( spanFreshJobs.LinePosition );

			}

			div = from lnks in doc.DocumentNode.Descendants()
				  where lnks.Name == "span" && lnks.InnerText.Contains( "Still Active" )
				  select new
				  {
					  spanOfInterest = lnks
				  };

			HtmlNode spanStillActive = null;
			if( div.Count() > 0 )
			{
				spanStillActive = div.First().spanOfInterest;
				//System.Diagnostics.Debug.WriteLine( spanStillActive.LinePosition );
			}



again:
			Int16 aboveFreshJobs = 0;
			Int16 belowStillActive = 0;
			Int16 jobCounter = 0;
			bool citiesFound = false;

			if( doc.DocumentNode.SelectNodes( "//a[@href]" ) != null )
				foreach( HtmlNode link in doc.DocumentNode.SelectNodes( "//a[@href]" ) )
				{
					if( !string.IsNullOrEmpty( link.InnerText ) )
					{
						foreach( Entity entity in this.cities )
						{
							try
							{
								if( link.InnerText.Contains( entity.Item ) )
								{
									citiesFound = true;
									string innerHtml = link.InnerHtml.Replace( entity.Item, string.Format( "<font color='COLORVALUE'>{0}</font>", entity.Item ) );
									System.Drawing.Color color =
										ColorInterpolator.InterpolateBetween( System.Drawing.Color.Green, System.Drawing.Color.Red, (double)entity.Attrib / 50 );
									innerHtml = innerHtml.Replace( "COLORVALUE", System.Drawing.ColorTranslator.ToHtml( color ) );
									link.InnerHtml = innerHtml;
									//
								}

							}
							catch( Exception e )
							{
							}
							//
						}
					}
				}

			if( !citiesFound )
			{
				if( doc.DocumentNode.SelectNodes( "//td" ) != null )
					foreach( HtmlNode link in doc.DocumentNode.SelectNodes( "//td" ) )
					{
						if( !string.IsNullOrEmpty( link.InnerText ) )
						{
							foreach( Entity entity in this.cities )
							{
								try
								{
									if( link.InnerText.Contains( entity.Item ) )
									{
										string innerHtml = link.InnerHtml.Replace( entity.Item, string.Format( "<font color='COLORVALUE'>{0}</font>", entity.Item ) );
										System.Drawing.Color color =
											ColorInterpolator.InterpolateBetween( System.Drawing.Color.Green, System.Drawing.Color.Red, (double)entity.Attrib / 50 );
										innerHtml = innerHtml.Replace( "COLORVALUE", System.Drawing.ColorTranslator.ToHtml( color ) );
										link.InnerHtml = innerHtml;
										//
									}

								}
								catch( Exception e )
								{
								}
								//
							}
						}
					}

			}

			if( doc.DocumentNode.SelectNodes( "//a[@href]" ) != null )
				foreach( HtmlNode link in doc.DocumentNode.SelectNodes( "//a[@href]" ) )
				{
					if( !string.IsNullOrEmpty( link.InnerText ) )
					{
						foreach( Entity entity in this.roles )
						{
							if( ( entity.Attrib == 2 ) && ( link.InnerText.Contains( entity.Item ) ) )
							{
								if( IsIrrelevant( link.InnerText.ToLower() ) )
									continue;

								bool previouslySubmitted = PreviouslySubmitted( link );

								jobCounter++;
								if( spanStillActive != null )
								{
									if( link.LinePosition < spanStillActive.LinePosition )
										aboveFreshJobs++;
									else if( link.LinePosition > spanStillActive.LinePosition )
										belowStillActive++;
									//
								}

								//#521987 = purple
								//#fffdaf = light yellow   </strike>
								if( previouslySubmitted )
									link.InnerHtml = link.InnerHtml.Replace( entity.Item, string.Format( "<span style='background-color:#fffdaf; color:purple'><strike><font size='5' >{0}</font></strike></span>", entity.Item ) );
								else
									link.InnerHtml = link.InnerHtml.Replace( entity.Item, string.Format( "<span style='background-color:#fffdaf; color:purple'><font size='5' >{0}</font></span>", entity.Item ) );

								//link.InnerHtml.Replace( link.InnerText, string.Format( "<h2 style='background-color:#fffdaf; color:#521987;'>{0}</h2>", link.InnerText ) );
								//link.InnerHtml.Replace( link.InnerText, string.Format( "<span style='background-color:red'><font color='green'>{0}</font></span>", link.InnerText ) );
								//link.InnerHtml.Replace( entity.Item, string.Format( "<span style='background-color:#fffdaf'><font color='blue'>{0}</font></span>", entity.Item ) );
								//link.InnerHtml.Replace( entity.Item, string.Format( "<font size='5' color='green'>{0}</font>", entity.Item ) );
								try
								{
									string pat = @"(\d+) ([a-zA-Z]+) ago";
									System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex( pat, System.Text.RegularExpressions.RegexOptions.IgnoreCase );
									System.Text.RegularExpressions.Match m = r.Match( link.InnerHtml );
									if( m.Success )
									{
										while( m.Success )
										{
											if( m.Groups.Count == 3 )
											{
												Int16 timeAgo = Convert.ToInt16( m.Groups[1].Value );
												string timePeriod = m.Groups[2].Value;

												link.InnerHtml = link.InnerHtml.Replace( m.Value, string.Format( "<font color='COLORVALUE'>{0}</font>", m.Value ) );

												if( timePeriod.Contains( "month" ) )
													timeAgo *= 30;
												else if( timePeriod.Contains( "week" ) )
													timeAgo *= 7;
												else if( timePeriod.Contains( "day" ) )
												{
													//timeAgo = timeAgo;
												}
												else if( timePeriod.Contains( "hour" ) )
													timeAgo = 1;
												else if( timePeriod.Contains( "second" ) )
													timeAgo = 1;

												System.Drawing.Color color = ColorInterpolator.InterpolateBetween( System.Drawing.Color.Green, System.Drawing.Color.Red, (double)timeAgo / 30 );
												//String strHtmlColor = System.Drawing.ColorTranslator.ToHtml(color);
												link.InnerHtml = link.InnerHtml.Replace( "COLORVALUE", System.Drawing.ColorTranslator.ToHtml( color ) );
												//
											}

											m = m.NextMatch();
											//
										}

									}//if (m.Success)
									 //
								}
								catch( Exception e )
								{
								}
								//
							}
							//
						}
						//
					}
					//
				}

			//
			//goto again;

			if( ( spanStillActive != null ) && ( aboveFreshJobs == 0 ) )
				throw new NoActiveJobsException();



			//if( spanFreshJobs != null )
			//	spanFreshJobs.InnerHtml = aboveFreshJobs + " Fresh Jobs";
			if( spanFreshJobs != null )
				spanFreshJobs.InnerHtml = "<span style='font - size:14px;color:red;font-family:Arial, sans-serif;'>" + aboveFreshJobs + " Fresh Jobs </span>";

			if( spanStillActive != null )
				spanStillActive.InnerHtml = "<span style='font - size:14px;color:red;font-family:Arial, sans-serif;'>" + belowStillActive + " Still Active </span>";
			//spanStillActive.InnerHtml = belowStillActive + " Still Active";


			return doc.DocumentNode.InnerHtml + string.Format( "<font size='1' color='red'><div id='gslprocesed' > Processed on: {0} Found {1}</div><font>", System.DateTime.Now.ToShortDateString(), jobCounter );
			//
		}

		WebClient webClient = new WebClient();
		private bool PreviouslySubmitted( HtmlNode link )
		{
again:

			string url = link.GetAttributeValue( "href", string.Empty );
			if( String.IsNullOrEmpty( url ) )
				return false;

			url = url.Split( '?' )[0];

			var getSubmittedJobsUrl = "http://localhost:56491/api/GetSubmittedJob?url=" + url;

			using( WebClient wc = new WebClient() )
			{
				var json = wc.DownloadString( getSubmittedJobsUrl );
zzz:
				try
				{
					JobResult output = (JobResult)Newtonsoft.Json.JsonConvert.DeserializeObject<JobResult>( json );
					if( output.statusCode == 404 )
						return false;

					return true;

				}
				catch( Exception ex )
				{
					//throw ex;
				}
				//

				//goto zzz;
			}

			//goto again;

			return true;
			//
		}


		public class JobResult
		{
			public int statusCode { get; set; }
			public string message { get; set; }
		}


		private bool IsIrrelevant( string text )
		{
			foreach( Entity entity in this.roles )
			{
				//System.Diagnostics.Debug.WriteLine( "{0} {1} {2}", text, entity.Item, ( text.Contains( entity.Item ) ) );
				if( ( entity.Attrib == 4 ) && ( text.Contains( entity.Item ) ) )
					return true;
			}

			return false;
			//
		}

		private void Inspectors_NewInspector( Outlook.Inspector Inspector )
		{
			Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
			if( mailItem != null )
			{
				if( mailItem.EntryID == null )
				{
					mailItem.Subject = "This text was added by using code";
					mailItem.Body = "This text was added by using code";
				}

			}

		}

		private void Read_Mails()
		{
			try
			{
				Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();//.ApplicationClass();

				Microsoft.Office.Interop.Outlook.NameSpace NS = app.GetNamespace( "MAPI" );

				Microsoft.Office.Interop.Outlook.MAPIFolder objFolder = NS.GetDefaultFolder( Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox );

				Microsoft.Office.Interop.Outlook.MailItem objMail;

				Microsoft.Office.Interop.Outlook.Items oItems;

				oItems = objFolder.Items;

				MessageBox.Show( "Reading mails" );

				for( int i = 1; i <= app.ActiveExplorer().Selection.Count; i++ )
				{

					MessageBox.Show( "Reading :" + i.ToString() );
					objMail = (Microsoft.Office.Interop.Outlook.MailItem)app.ActiveExplorer().Selection[i];
					MessageBox.Show( objMail.Body.ToString() );
				}



			}
			catch( Exception ex )
			{
				MessageBox.Show( ex.ToString() );
			}

			finally
			{
				//NS.Logoff();

				//objFolder = null;

				//objMail = null;

				//app = null;
			}

		}

		//protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
		//{
		//	return new JobRibbon();
		//}


		private void ThisAddIn_Shutdown( object sender, System.EventArgs e )
		{
			// Note: Outlook no longer raises this event. If you have code that 
			//    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler( ThisAddIn_Startup );
			this.Shutdown += new System.EventHandler( ThisAddIn_Shutdown );
		}

		[Serializable]
		private class NoActiveJobsException : Exception
		{
			public NoActiveJobsException()
			{
			}

			public NoActiveJobsException( string message ) : base( message )
			{
			}

			public NoActiveJobsException( string message, Exception innerException ) : base( message, innerException )
			{
			}

			protected NoActiveJobsException( SerializationInfo info, StreamingContext context ) : base( info, context )
			{
			}
		}

		#endregion
	}
}
