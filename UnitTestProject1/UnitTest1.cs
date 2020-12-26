using System;
using System.IO;
using EmailProcessor;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        Microsoft.Office.Interop.Outlook.Application app;
        Microsoft.Office.Interop.Word.Application workApp;
        EmailProcessor.EmailProcessor emailProcessor = null;
        string content = null;

        public UnitTest1()
        {
            app = new Microsoft.Office.Interop.Outlook.Application();

            emailProcessor = new EmailProcessorSportery( app );
            string fileName = "./bookedEmail.txt";
            if( !File.Exists( fileName ) )
            {
                Assert.Fail( fileName + " not found" );
                return;
            }

            StreamReader sr = File.OpenText( fileName );
            content = sr.ReadToEnd();
            //
        }
        [TestMethod]

        public void Test1()
        {
        abc:
            try
            {
                //emailProcessor.DeleteCalendarItem( content );
                AppointmentItem newAppointment = (AppointmentItem)emailProcessor.CreateCalendarItem( content );
                if( newAppointment != null )
                {
                    newAppointment.Display();
                    //newAppointment.Save();
                }
                //Assert.Pass();
                //
            }
            catch( System.Exception ex )
            {
                //TestContext.Out.WriteLine( ex.Message );
                //TestContext.Out.WriteLine( ex.StackTrace );
            }

            //TestContext.Error.WriteLine( "fuck " );
            //TestContext.WriteLine( " off" );
            //
            //goto abc;
        }
        //


    }
}
