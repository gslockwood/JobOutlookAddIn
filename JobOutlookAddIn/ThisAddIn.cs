using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Utilities;
using System.Runtime.Serialization;
using System.Net;
using Newtonsoft.Json;
using HtmlAgilityPack;
using System.Threading.Tasks;
using EmailProcessor;

namespace JobOutlookAddIn
{
    public partial class ThisAddIn
    {
        public static string version = " v1.41";
        //public static Outlook.Application app;

        WebClient webClient = new WebClient();
        private Outlook.NameSpace outlookNameSpace;

        Outlook.Inspectors inspectors;
        Outlook.Explorer currentExplorer = null;
        private Dictionary<int, IList<Entity>> dict;
        IList<Entity> roles;
        IList<Entity> cities;
        JobEmailResponce jobEmailResponce = null;

        //reservationsspecialtennis@gmail.com
        private readonly string gslprocesed = "gslprocesed";
        private bool testing = false;//true false
        private string lastEntryID;

        private void ThisAddIn_Startup( object sender, System.EventArgs e )
        {
            inspectors = this.Application.Inspectors;
            //inspectors.NewInspector += Inspectors_NewInspector;

            //app = this.Application;

            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += CurrentExplorer_SelectionChange;

            outlookNameSpace = this.Application.GetNamespace( "MAPI" );
            /*
			Outlook.MAPIFolder inbox = outlookNameSpace.GetDefaultFolder( Outlook.OlDefaultFolders.olFolderInbox );
			inbox.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler( items_ItemAdd );
			//outlookNameSpace.
			*/
            Application.NewMailEx += Application_NewMailEx;

            ProcessTextfiles();
            //
        }

        private void Application_NewMailEx( string entryIDCollection )
        {
            try
            {
                dynamic newITem = outlookNameSpace.GetItemFromID( entryIDCollection );
                Outlook.MailItem mailItem = (Outlook.MailItem)newITem;
                if( mailItem != null )
                {
                    ProcessNewMail( mailItem );
                }
                //
            }
            catch( Exception ex )
            {
                //throw ex;
            }
            //
            return;
            //
        }

        private Outlook.Items GetAppointmentsInRange( DateTime startTime, DateTime endTime, string filterString )
        {
            Outlook.Folder calFolder = Application.Session.GetDefaultFolder( Outlook.OlDefaultFolders.olFolderCalendar ) as Outlook.Folder;

            string filter = "[Start] >= '"
                + startTime.ToString( "g" )
                + "' AND [End] < '"
                + endTime.ToString( "g" ) + "'";
            //Console.WriteLine( filter );

            try
            {
                Outlook.Items calItems = calFolder.Items;
                Outlook.Items restrictItems = calItems.Restrict( filter );
                restrictItems = restrictItems.Restrict( filterString );
                if( restrictItems.Count > 0 )
                    return restrictItems;
                else
                    return null;
            }
            catch( Exception ex )
            {
                return null;
            }
            //
        }

        private void ProcessNewMail( Outlook.MailItem mailItem )
        {
            var senderEmailAddress = mailItem.SenderEmailAddress;
            var htmlBody = mailItem.HTMLBody;
            htmlBody = System.Net.WebUtility.HtmlDecode( htmlBody );

            EmailProcessor.EmailProcessor emailProcessor = null;
            if( senderEmailAddress == "support@spotery.com" )
            {
                emailProcessor = new EmailProcessorSportery( this.Application );

                if( mailItem.Subject.Contains( "updated" ) )
                {
                    if( mailItem.Body.Contains( "Canceled by User" ) )
                    {
                        emailProcessor.DeleteCalendarItem( htmlBody );
                    }

                }
                else if( mailItem.Subject.Contains( "confirmed" ) )
                {
                    Outlook.AppointmentItem newAppointment = (Outlook.AppointmentItem)emailProcessor.CreateCalendarItem( htmlBody );

                    if( newAppointment != null )
                    {
                        //newAppointment.Display();
                        newAppointment.Save();
                    }

                }
                //
            }

            else if( senderEmailAddress == "reservationsspecialtennis@gmail.com" )
            {
                try
                {
                    //html += "<section id='ReservationResponce' hidden>";
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml( htmlBody );
                    HtmlNode node = doc.GetElementbyId( "ReservationResponce" );
                    if( node != null )
                    {
                        string content = node.InnerText;
                        System.Xml.XmlDocument xNode = JsonConvert.DeserializeXmlNode( "{ \"ReservationResponce\":" + content + "}" );
                        GGTSReservations.Services.ReservationResponce reservation = JsonConvert.DeserializeObject<GGTSReservations.Services.ReservationResponce>( content );

                        Outlook.Items appts = GetAppointmentsInRange( Convert.ToDateTime( reservation.court_date ), Convert.ToDateTime( reservation.court_date ).AddDays( +1 ), "[BayClubReservation]='BayClubReservation'" );
                        if( ( appts != null ) && ( appts.Count > 0 ) )
                            return;

                        Outlook.AppointmentItem newAppointment = (Outlook.AppointmentItem)this.Application.CreateItem( Outlook.OlItemType.olAppointmentItem );
                        newAppointment.Start = Convert.ToDateTime( reservation.court_date + " " + reservation.start_time );
                        newAppointment.End = Convert.ToDateTime( reservation.court_date + " " + reservation.end_time ).AddHours( +1 );
                        newAppointment.ReminderMinutesBeforeStart = 90;
                        if( reservation.court_surface == "Gateway" )
                            newAppointment.Location = "370 Drumm St, San Francisco, CA 94111";

                        System.Text.StringBuilder sb = new System.Text.StringBuilder();
                        sb.AppendLine( string.Format( "Your Reservation at {0}, {1}", reservation.court_surface, newAppointment.Location ) );
                        sb.AppendLine( string.Format( "Your Reservation is set for {0} for {1}", newAppointment.Start.ToLongDateString(), reservation.court_name ) );
                        sb.AppendLine();
                        sb.AppendLine();
                        sb.AppendLine( "Players: " );
                        sb.AppendLine( string.Format( "\t{0} {1}", reservation.player_1_first_name, reservation.player_1_name ) );
                        if( reservation.player_2_id != 0 )
                            sb.AppendLine( string.Format( "\t{0} {1}", reservation.player_2_first_name, reservation.player_2_name ) );
                        if( reservation.player_3_id != 0 )
                            sb.AppendLine( string.Format( "\t{0} {1}", reservation.player_3_first_name, reservation.player_3_name ) );
                        if( reservation.player_4_id != 0 )
                            sb.AppendLine( string.Format( "\t{0} {1}", reservation.player_4_first_name, reservation.player_4_name ) );

                        sb.AppendLine();
                        sb.AppendLine();
                        sb.AppendLine( string.Format( "Created on: {0}", DateTime.Now.ToString( "MMM ddd d HH:mm yyyy" ) ) );

                        newAppointment.Body = sb.ToString();

                        newAppointment.Subject = "Auto Tennis appt: Club: " + reservation.court_surface + " on " + reservation.court_name + " at " + newAppointment.Start.ToShortTimeString();


                        Outlook.ItemProperty MeetingNameProperty = newAppointment.ItemProperties.Add( "BayClubReservation", Outlook.OlUserPropertyType.olText, true );
                        MeetingNameProperty.Value = "BayClubReservation";

                        newAppointment.Save();
                        //
                    }
                    /*
					 {"ball_machine": 1,"booking_time": "2018-07-31 08:23:25","court_blocks": "[\"a_034_09_00_2018-08-02_8\", \"a_034_09_30_2018-08-02_9\"]","court_date": "2018-08-02","court_id": 1043910,"court_length": 1.000000,"court_name": "Tennis 6 Ball Machine","court_number": 34,"court_range_id": 4,"court_sport": "Tennis","court_status": 0,"court_surface": "Gateway","court_uid": "3a2b9004-cbdc-45c7-b7f7-fe32cae12137","end_time": "10:00:00","errormap": {"rule_0": "fail"},"mode": 30,"number_of_players": 1,"online_booking": 1,"player_1_first_name": "George","player_1_id": 158845,"player_1_name": "Lockwood","player_2_first_name": "","player_2_id": 0,"player_2_name": "","player_3_first_name": "","player_3_id": 0,"player_3_name": "","player_4_first_name": "","player_4_id": 0,"player_4_name": "","reservation_type": "Open","result": "OK","start_time": "09:00:00","successmap": {"rule_0": "success","rule_1": "Allowed to view courts","rule_12": "Booking Allowed (Access)","rule_13": "has not exceeded 1.5 hr. Currently have 0.000000 hr.1.000000","rule_16": "No side by side","rule_2": "Booking time reached","rule_3": "Booking in future","rule_5": "No back to back booking","rule_6": "No back to back booking (2nd Player)","rule_9": "Tennis 6 Ball Machine can be booked online"},"view_days": 7}
					 */
                }
                catch( Exception ex )
                {
                    //throw ex;
                }
                //
            }
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
        again:
            try
            {
                string userDir = Environment.GetFolderPath( Environment.SpecialFolder.UserProfile );
                //C:\Users\Georg\OneDrive\OutlookHelper
                dict = FileUtilities.ParseFileEx( userDir + @"\OneDrive\OutlookHelper\Roles.txt" );
                roles = FileUtilities.ParseFile( userDir + @"\OneDrive\OutlookHelper\Roles.txt" );
                cities = FileUtilities.ParseFile( userDir + @"\OneDrive\OutlookHelper\Cities.txt" );

                string resumeFilename = userDir + @"\OneDrive\Resumes\vp.docx";
                string tempFilename = userDir + @"\OneDrive\Resumes\georgelockwoodresume.docx";

                System.IO.FileAttributes attribute = System.IO.FileAttributes.ReadOnly & System.IO.FileAttributes.Archive;

                System.IO.File.SetAttributes( tempFilename, ~System.IO.FileAttributes.ReadOnly );
                System.IO.File.Delete( tempFilename );
                /*
				if( !System.IO.File.Exists( tempFilename ) )
					throw new Exception( string.Format( "the {0} file is missing", resumeFilename ) );
				*/

                System.IO.File.Copy( resumeFilename, tempFilename, true );

                //System.IO.File.SetAttributes( tempFilename, attribute );
                //System.IO.File.SetAttributes( tempFilename, attribute );

                //jobEmailResponce = new JobEmailResponce( userDir + @"\OneDrive\OutlookHelper\outgoingmessage.txt", userDir + @"\OneDrive\Resumes\vp.docx", roles );
                jobEmailResponce = new JobEmailResponce( userDir + @"\OneDrive\OutlookHelper\outgoingmessage.txt", tempFilename, roles );
                //
            }
            catch( Exception x )
            {
                MessageBox.Show( x.Message );
            }
            //goto again;
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

                        if( mailItem.EntryID == lastEntryID ) return;

                        lastEntryID = mailItem.EntryID;

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

        private async void ProcessEmail( Outlook.MailItem mailItem )
        {
            string htmlBody = mailItem.HTMLBody;
            try
            {
                /////////////////////testing? or keep?
                ProcessNewMail( mailItem );

                //var temp = mailItem.HTMLBody;

                var result = await ProcessBodyAsync( mailItem.SenderEmailAddress, htmlBody );
                if( ( result != null ) )//&& ( result.Result != null ) )
                    mailItem.HTMLBody = result;//.Result;

                //mailItem.HTMLBody = temp;

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

        private async Task<string> ProcessBodyAsync( string address, string htmlBody )
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

            if( !testing )
                if( doc.GetElementbyId( gslprocesed ) != null )
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

            IList<HtmlNode> nodes = doc.DocumentNode.SelectNodes( "//a[@href]" );
            if( nodes != null )
            {
                foreach( HtmlNode link in nodes )
                {
                    if( string.IsNullOrEmpty( link.InnerText.Trim() ) )
                        continue;

                    //link.InnerHtml = link.InnerHtml.Replace( "Irrelevant ", "" );
                    string temp = link.InnerText;
                    if( temp.Length > 35 )
                        temp = link.InnerText.Substring( 0, 36 );

                    if( IsIrrelevantEx( temp.ToLower() ) )
                    {
                        link.InnerHtml = "<span style='background-color:lightgrey; color:red'>Irrelevant </span>" + link.InnerHtml;
                        continue;
                    }

                    bool found = false;

                    //foreach( Entity entity in this.roles )
                    foreach( Entity entity in this.dict[2] )
                    {
                        //if( ( entity.Attrib == 2 ) && ( link.InnerText.Contains( entity.Item ) ) )
                        if( link.InnerText.Contains( entity.Item ) )
                        {
                            found = true;

                            jobCounter++;
                            if( spanStillActive != null )
                            {
                                if( link.LinePosition < spanStillActive.LinePosition )
                                    aboveFreshJobs++;
                                else if( link.LinePosition > spanStillActive.LinePosition )
                                    belowStillActive++;
                                //
                            }

                            /*  this is never finding jobs in the DB maybe search on the jobid only?
                            bool previouslySubmitted = await PreviouslySubmitted( link );

                            //#521987 = purple
                            //#fffdaf = light yellow   </strike>
                            if( previouslySubmitted )
                                link.InnerHtml = link.InnerHtml.Replace( entity.Item, string.Format( "<span style='background-color:#fffdaf; color:purple'><strike><font size='5' >Submitted: {0}</font></strike></span>", entity.Item ) );
                            else
                                link.InnerHtml = link.InnerHtml.Replace( entity.Item, string.Format( "<span style='background-color:#fffdaf; color:purple'><font size='5' >{0}</font></span>", entity.Item ) );
                            */

                            link.InnerHtml =
                                link.InnerHtml.Replace( entity.Item, string.Format( "<span style='background-color:#fffdaf; color:lightblue'><font size='5' >{0}</font></span>", entity.Item ) );

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

                            break;
                            //
                        }//if( link.InnerText.Contains( entity.Item ) )
                        //
                    }//foreach( Entity entity in this.dict[2] )

                    if( found )
                        continue;

                    foreach( Entity entity in this.dict[0] )
                    {
                        if( link.InnerText.Contains( entity.Item ) )
                        {
                            found = true;
                            link.InnerHtml = link.InnerHtml.Replace( entity.Item, string.Format( "<span style='background-color:lightgrey; color:DarkRed'><font size='5' ><strike>{0}</strike></font></span>", entity.Item ) );
                            break;
                        }
                    }

                    if( found )
                        continue;

                    foreach( Entity entity in this.dict[1] )
                    {
                        if( link.InnerText.Contains( entity.Item ) )
                        {
                            found = true;
                            link.InnerHtml = link.InnerHtml.Replace( entity.Item, string.Format( "<span style='background-color:lightgrey; color:DarkRed'><font size='5' ><strike>{0}</strike></font></span>", entity.Item ) );
                            break;
                        }
                    }
                    //
                }//foreach( HtmlNode link in doc.DocumentNode.SelectNodes( "//a[@href]" ) )

            }

            if( ( spanStillActive != null ) && ( aboveFreshJobs == 0 ) )
                throw new NoActiveJobsException();

            if( ( htmlBody.Contains( "Jobsite" ) ) && ( jobCounter == 0 ) )
                throw new NoActiveJobsException();
            // 
            if( ( address.Contains( "cityjobs@cityjobsmail.com" ) ) && ( jobCounter == 0 ) )
                throw new NoActiveJobsException();

            if( ( address.Contains( "lensa" ) ) && ( jobCounter == 0 ) )
                throw new NoActiveJobsException();

            if( ( address.Contains( "talent@angel.co" ) ) && ( htmlBody.Contains( "New job listings:" ) ) && ( jobCounter == 0 ) )
                throw new NoActiveJobsException();



            //if( spanFreshJobs != null )
            //	spanFreshJobs.InnerHtml = aboveFreshJobs + " Fresh Jobs";
            if( spanFreshJobs != null )
                spanFreshJobs.InnerHtml = "<span style='font - size:14px;color:red;font-family:Arial, sans-serif;'>" + aboveFreshJobs + " Fresh Jobs </span>";

            if( spanStillActive != null )
                spanStillActive.InnerHtml = "<span style='font - size:14px;color:red;font-family:Arial, sans-serif;'>" + belowStillActive + " Still Active </span>";
            //spanStillActive.InnerHtml = belowStillActive + " Still Active";


            return doc.DocumentNode.InnerHtml + string.Format( "<font size='1' color='red'><div id='" + gslprocesed + "' > Processed on: {0} Found {1}</div><font>", System.DateTime.Now.ToShortDateString(), jobCounter );
            //return doc.DocumentNode.InnerHtml + string.Format( "<font size='1' color='red'><div id='gslprocesed' > Processed on: {0} Found {1}</div><font>", System.DateTime.Now.ToShortDateString(), jobCounter );
            //
        }

        public class RyadelWebClient : WebClient
        {
            /// <summary>
            /// Default constructor (30000 ms timeout)
            /// NOTE: timeout can be changed later on using the [Timeout] property.
            /// </summary>
            public RyadelWebClient() : this( 10000 ) { }

            /// <summary>
            /// Constructor with customizable timeout
            /// </summary>
            /// <param name="timeout">
            /// Web request timeout (in milliseconds)
            /// </param>
            public RyadelWebClient( int timeout )
            {
                Timeout = timeout;
            }

            #region Methods
            protected override WebRequest GetWebRequest( Uri uri )
            {
                WebRequest w = base.GetWebRequest( uri );
                w.Timeout = Timeout;
                ( (HttpWebRequest)w ).ReadWriteTimeout = Timeout;
                return w;
            }

            public new async Task<string> DownloadStringTaskAsync( Uri address )
            {
                var t = base.DownloadStringTaskAsync( address );
                if( await Task.WhenAny( t, Task.Delay( Timeout ) ) != t )
                    CancelAsync();
                return await t;
            }
            public new async Task<string> DownloadStringTaskAsync( string address )
            {
                var t = base.DownloadStringTaskAsync( address );
                if( await Task.WhenAny( t, Task.Delay( Timeout ) ) != t )
                    CancelAsync();
                return await t;
            }
            #endregion

            /// <summary>
            /// Web request timeout (in milliseconds)
            /// </summary>
            public int Timeout { get; set; }
        }

        private async Task<bool> PreviouslySubmitted( HtmlNode link )
        {
            string url = link.GetAttributeValue( "href", string.Empty );
            if( String.IsNullOrEmpty( url ) )
                return false;

            url = url.Split( '?' )[0];

            var getSubmittedJobsUrl = "http://localhost:56491/api/GetSubmittedJob?url=" + url;

            //using( WebClient wc = new WebClient() )
            using( RyadelWebClient wc = new RyadelWebClient() )
            {
                //var json = wc.DownloadString( getSubmittedJobsUrl );
                var json = await wc.DownloadStringTaskAsync( getSubmittedJobsUrl );
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
            }

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
        private bool IsIrrelevantEx( string text )
        {
            IList<Entity> roles4 = this.dict[4];
            foreach( Entity entity in roles4 )
            {
                //System.Diagnostics.Debug.WriteLine( "{0} {1} {2}", text, entity.Item, ( text.Contains( entity.Item ) ) );
                if( text.Contains( entity.Item ) )
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

        public enum FileCategory
        {
            Role,
            City,
            OutGoingMessage
        }

        internal void EditFile( FileCategory fileCategory )
        {
            string path = @"C:\Users\Georg\OneDrive\OutlookHelper\";
            if( fileCategory == FileCategory.Role )
                path += "Roles.txt";
            else if( fileCategory == FileCategory.City )
                path += "Cities.txt";
            else if( fileCategory == FileCategory.OutGoingMessage )
                path += "OutGoingMessage.txt";

            System.Diagnostics.Process.Start( path );
            //
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


namespace GGTSReservations.Services
{
    public class ReservationResponce
    {
        //[Key]
        public int Id { get; set; }
        public int ball_machine { get; set; }
        public string booking_time { get; set; }
        public string court_blocks { get; set; }
        public string court_date { get; set; }
        public int court_id { get; set; }
        public string court_uid { get; set; }

        public string court_name { get; set; }
        public string court_sport { get; set; }
        public string court_surface { get; set; }

        public float court_length { get; set; }
        public int court_number { get; set; }
        public int court_range_id { get; set; }
        public int court_status { get; set; }
        public string end_time { get; set; }
        public Errormap errormap { get; set; }
        public int mode { get; set; }
        public int number_of_players { get; set; }
        public int online_booking { get; set; }
        public string player_1_first_name { get; set; }
        public int player_1_id { get; set; }
        public string player_1_name { get; set; }
        public string player_2_first_name { get; set; }
        public int player_2_id { get; set; }
        public string player_2_name { get; set; }
        public string player_3_first_name { get; set; }
        public int player_3_id { get; set; }
        public string player_3_name { get; set; }
        public string player_4_first_name { get; set; }
        public int player_4_id { get; set; }
        public string player_4_name { get; set; }
        public string reservation_type { get; set; }
        public string result { get; set; }
        public string start_time { get; set; }
        public Successmap successmap { get; set; }
        public int view_days { get; set; }

        // added
        public string eventId { get; set; }
        public bool Unavailable { get; set; }
        //
    }

    public class Errormap
    {
        //private int count = 1;
        //public int count { get; set; }

        public string rule_0 { get; set; }
        public string rule_13 { get; set; }
        public string rule_16 { get; set; }
        public string rule_5 { get; set; }
        public string rule_12 { get; set; }
        public string rule_3 { get; set; }
        public string rule_2 { get; set; }

        public int GetPropertiesCount()
        {
            int counter = 0;
            foreach( System.Reflection.PropertyInfo property in this.GetType().GetProperties() )
            {
                if( property.Name.Equals( "rule_0" ) )
                    continue;

                string value = (string)property.GetValue( this );
                //Console.WriteLine( value.Length );
                if( value != null )
                    counter++;

            }
            //goto gggg;
            //Console.WriteLine( counter );

            return counter;

            //return this.GetType().GetProperties().Count();

        }

        //public int GetPropertiesCount()
        //{
        //    return count;

        //}
    }

    public class Successmap
    {
        public string rule_0 { get; set; }
        public string rule_1 { get; set; }
        public string rule_12 { get; set; }
        public string rule_2 { get; set; }
        public string rule_3 { get; set; }
        public string rule_6 { get; set; }
        public string rule_9 { get; set; }
    }

}
