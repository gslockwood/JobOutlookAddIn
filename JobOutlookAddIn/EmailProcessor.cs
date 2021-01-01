using JobOutlookAddIn;
using Microsoft.Office.Interop.Outlook;
//using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EmailProcessor
{
    public class EmailProcessor : ICreateCalendarItem
    {
        protected Microsoft.Office.Interop.Outlook.Application app;
        protected string ReservationUserPropertyTitle = "baseReservation";
        protected string ReservationUserPropertyValue = "Reservation:";
        protected string senderEmailAddress;

        public EmailProcessor( Microsoft.Office.Interop.Outlook.Application app )
        {
            this.app = app;
        }

        public virtual object CreateCalendarItem( string content )
        {
            return new NotImplementedException();
        }

        public virtual void DeleteCalendarItem( string content )
        {
            throw new NotImplementedException();
        }


        /*
        public virtual void DeleteCalendarItem( string reservationNumber )
        {
            Items items = GetAppointmentsByReservationNumber( reservationNumber );
            if( ( items != null ) && ( items.Count > 0 ) )
            {
                AppointmentItem itemDelete = items.GetFirst();
                itemDelete.Delete();
            }

        }
        */

        protected Items GetAppointmentsByReservationNumber( string reservationNumber )
        {
            Folder calFolder = app.Session.GetDefaultFolder( OlDefaultFolders.olFolderCalendar ) as Folder;
            try
            {
                Items calItems = calFolder.Items;
                //                                   "[BayClubReservation]='Reservation'" );
                string filterString = string.Format( "[{0}]='{1}{2}'", this.ReservationUserPropertyTitle, this.ReservationUserPropertyValue, reservationNumber );
                Items restrictItems = calItems.Restrict( filterString );

                return ( restrictItems.Count > 0 ) ? restrictItems : null;
                //
            }
            catch( System.Exception ex )
            {
                return null;
            }
            //
        }

        public virtual void ProcessMail( MailItem mailItem )
        {
            throw new NotImplementedException();
        }
        //
    }

    public class EmailProcessorSportery : EmailProcessor, ICreateCalendarItem
    {
        public EmailProcessorSportery( Microsoft.Office.Interop.Outlook.Application app ) : base( app )
        {
            ReservationUserPropertyTitle = "SpoteryReservation";
        }

        public override void ProcessMail( MailItem mailItem )
        {
            senderEmailAddress = mailItem.SenderEmailAddress;
            if( mailItem.Subject.Contains( "updated" ) )
            {
                if( mailItem.Body.Contains( "Canceled by User" ) )
                {
                    DeleteCalendarItem( mailItem.Body );
                }

            }
            else if( mailItem.Subject.Contains( "confirmed" ) )
            {
                AppointmentItem newAppointment = (AppointmentItem)CreateCalendarItem( mailItem.Body );

                if( newAppointment != null )
                {
                    //newAppointment.Display();
                    newAppointment.Save();
                }

            }
            //
        }

        public override void DeleteCalendarItem( string content )
        {
            string reservationsNumber = "ReservationsNumber not found";
            //<div><strong>Reservation #:</strong> 1322312<br>
            string pattern = "<strong>Reservation #:</strong> (\\d+)<br>";
            Regex regex = new Regex( pattern );
            Match match = regex.Match( content );
            if( ( match.Success ) && ( match.Groups.Count > 1 ) )
            {
                reservationsNumber = " #" + match.Groups[1].Value;

                Items items = GetAppointmentsByReservationNumber( reservationsNumber );
                if( ( items != null ) && ( items.Count > 0 ) )
                    items.GetFirst().Delete();
                
                return;
                //
            }

            pattern = "(Reservation #: (\\d+))";
            regex = new Regex( pattern );
            match = regex.Match( content );
            if( ( match.Success ) && ( match.Groups.Count > 1 ) )
            {
                reservationsNumber = " #" + match.Groups[2].Value;
                Items items = GetAppointmentsByReservationNumber( reservationsNumber );
                if( ( items != null ) && ( items.Count > 0 ) )
                    items.GetFirst().Delete();

                return;
                //
            }

        }

        bool testing = false;//false true
        public override object CreateCalendarItem( string content )
        {
        //workApp = new Microsoft.Office.Interop.Word.Application();
        abc:
            string reservationsNumber = "ReservationsNumber not found";
            //<div><strong>Reservation #:</strong> 1322312<br>
            string pattern = "<strong>Reservation #:</strong> (.*)<br>";
            Regex regex = new Regex( pattern );
            Match match = regex.Match( content );
            if( ( match.Success ) && ( match.Groups.Count > 1 ) )
            {
                reservationsNumber = " #" + match.Groups[1].Value;
                //newAppointment.UserProperties
                Items appts = GetAppointmentsByReservationNumber( reservationsNumber );
                if( !testing )
                    if( ( appts != null ) && ( appts.Count > 0 ) )
                        return null;
            }

            /*
            //This email was sent to greenflashtennis@gmail.com
            string footer = "An email address was not found";
            //<strong>Spot Name:</strong> Lafayette Tennis Court #2</div>
            pattern = "(This email was sent to (.*))";
            regex = new Regex( pattern );
            match = regex.Match( content );
            if( ( match.Success ) && ( match.Groups.Count > 1 ) )
            {
                footer = match.Groups[1].Value;
                //footer = footer.TrimEnd( '\r' );
            }
            */

            //footer = "<div><div><div>" + footer.TrimEnd( '\r' ) + "</div></div></div>";          
            string footer = "<div><div><div>" + this.senderEmailAddress + "</div></div></div>";

            string location = "Location not found";
            //<strong>Spot Name:</strong> Lafayette Tennis Court #2</div>
            pattern = "<strong>Spot Name:</strong> (.*)</div>";
            regex = new Regex( pattern );
            match = regex.Match( content );
            if( ( match.Success ) && ( match.Groups.Count > 1 ) )
            {
                location = match.Groups[1].Value;
            }

            string body = "Body not found";

            string temp = content;

            //<div>Dear 
            int start = temp.IndexOf( "<div>Dear " );// - "<div>Dear ".Length;
            int end = temp.IndexOf( "</td>", start ) + "</td>".Length;
            int length = end - start;
            temp = temp.Substring( start, length );
            temp = temp.Replace( "\r\n", "" );
            temp = temp.Replace( "\t", "" );

            pattern = "(<div>Dear (.*))</td>";

            regex = new Regex( pattern );
            match = regex.Match( temp );
            if( ( match.Success ) && ( match.Groups.Count > 1 ) )
            {
                body = match.Groups[1].Value + "</div></div></div></div></div>";
            }

            AppointmentItem newAppointment = (AppointmentItem)app.CreateItem( OlItemType.olAppointmentItem );


            pattern = "<strong>Activity Date:</strong> (.*)</div>";
            regex = new Regex( pattern );
            match = regex.Match( content );
            if( ( match.Success ) && ( match.Groups.Count > 1 ) )
            {
                string value = match.Groups[1].Value;
                var array = value.Split( new string[] { " to " }, StringSplitOptions.None );
                try
                {
                    newAppointment.Start = DateTime.Parse( array[0] );
                    newAppointment.End = DateTime.Parse( newAppointment.Start.ToShortDateString() + " " + array[1] );
                }
                catch( System.Exception ex )
                {
                    //throw ex;
                }
                //
            }

            body += footer;

            newAppointment.Subject = string.Format( "{0}{1} ({2} minutes)", location, reservationsNumber, newAppointment.Duration );

            HTML2RTFConverter html2RTFConverter = new HTML2RTFConverter();
            string newTemp = html2RTFConverter.Convert( body );

            newAppointment.RTFBody = Encoding.ASCII.GetBytes( newTemp );

            //ItemProperty MeetingNameProperty = newAppointment.ItemProperties.Add( "SpoteryReservation", OlUserPropertyType.olText, true );
            ItemProperty MeetingNameProperty = newAppointment.ItemProperties.Add( ReservationUserPropertyTitle, OlUserPropertyType.olText, true );
            MeetingNameProperty.Value = this.ReservationUserPropertyValue + reservationsNumber;

            //workApp.Quit();

            return newAppointment;
            //
        }


    }
    public class EmailProcessorGGTS : EmailProcessor, ICreateCalendarItem
    {
        public EmailProcessorGGTS( Microsoft.Office.Interop.Outlook.Application app ) : base( app )
        {
            ReservationUserPropertyTitle = "GGTSReservation";
        }

        public override object CreateCalendarItem( string content )
        {
            return base.CreateCalendarItem( content );
        }

        public override void DeleteCalendarItem( string content )
        {
            base.DeleteCalendarItem( content );
        }

        public override void ProcessMail( MailItem mailItem )
        {
            base.ProcessMail( mailItem );
        }

    }
}
