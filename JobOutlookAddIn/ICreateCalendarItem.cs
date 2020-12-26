namespace EmailProcessor
{
    public interface ICreateCalendarItem
    {
        object CreateCalendarItem( string content );
        void DeleteCalendarItem( string content );
        //Microsoft.Office.Interop.Outlook.Items GetAppointmentsByFilterString( string reservationNumber );
    }
}