namespace EmailProcessor
{
    public interface ICreateCalendarItem
    {
        void ProcessMail( Microsoft.Office.Interop.Outlook.MailItem mailItem );
        object CreateCalendarItem( string content );
        void DeleteCalendarItem( string content );
    }
}