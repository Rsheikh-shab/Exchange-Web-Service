using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Configuration;

namespace ExchangeServices
{
    class Program
    {
        static ExchangeService service;
        static string userEmail = ConfigurationSettings.AppSettings["UserEmail"]; //"Email"
        static string userPassword = ConfigurationSettings.AppSettings["UserPassword"]; //"Password";
        static string serviceDomain = ConfigurationSettings.AppSettings["DomainName"]; //"Domain" if needed;
        static string serviceURL = ConfigurationSettings.AppSettings["Service Url"];

        static void Main(string[] args)
        {
            Console.WriteLine("Connecting to Exchange Online, please wait...");
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            service.Credentials = new WebCredentials(userEmail, userPassword);
            //NetworkCredential credential = new NetworkCredential(userEmail, userPassword, serviceDomain);
            service.AutodiscoverUrl(userEmail, RedirectionUrlValidationCallback);
            ExchangeService();
        }

        static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool redirectionValidated = false;
            if (redirectionUrl.Equals("https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml"))
                redirectionValidated = true;

            else if (redirectionUrl.Equals(serviceURL))
                redirectionValidated = true;

            return redirectionValidated;
        }

        static void DisplayMessage(string message)
        {
            Console.Clear();
            Console.WriteLine("\t\t\t\t************************************************");
            Console.WriteLine(message);
            Console.WriteLine("\t\t\t\t************************************************");
        }

        static void ExchangeService()
        {
            Console.Clear();

            Console.WriteLine("\t\t\t***************************************************************************");
            Console.WriteLine("\t\t\t*                                                                         *");
            Console.WriteLine("\t\t\t*                    Exchange Web Service                                 *");
            Console.WriteLine("\t\t\t*                                                                         *");
            Console.WriteLine("\t\t\t*                    What would you like to do?                           *");
            Console.WriteLine("\t\t\t*                                                                         *");
            Console.WriteLine("\t\t\t*                        Appointments  :                                  *");
            Console.WriteLine("\t\t\t*                       ----------------                                  *");
            Console.WriteLine("\t\t\t*                       (0) Exit                                          *");
            Console.WriteLine("\t\t\t*                       (1) Create Appointment                            *");
            Console.WriteLine("\t\t\t*                       (2) Cancel Appointment                            *");
            Console.WriteLine("\t\t\t*                       (3) Find All Appointment                          *");
            Console.WriteLine("\t\t\t*                       (4) Check Availability (Next 24 hours)            *");
            Console.WriteLine("\t\t\t*                                                                         *");
            Console.WriteLine("\t\t\t*                                                                         *");
            Console.WriteLine("\t\t\t***************************************************************************");

            Console.Write("\n\tEnter Your Choice: ");
            int choice = Convert.ToInt32(Console.ReadLine());

            switch (choice)
            {
                case 1:
                    CreateAppointment();
                    Console.ReadLine();
                    break;

                case 2:
                    CancelAppointment();
                    Console.ReadLine();
                    break;

                case 3:
                    FindAllAppointment();
                    Console.ReadLine();
                    break;

                case 4:
                    CheckAvailability();
                    Console.ReadLine();
                    break;

                case 0:
                    Environment.Exit(0);
                    break;

                default:
                    break;
            }
            ExchangeService();
            Console.ReadLine();
        }

        static void CreateAppointment()
        {
            DisplayMessage("\t\t\t\t\tCreating Appointment, please wait...");
            Appointment appointment = new Appointment(service);
            appointment.Subject = "Status Meeting";
            appointment.Body = "The purpose of this meeting is to discuss status";
            appointment.Location = "MM Level 15-1";
            appointment.Start = DateTime.Now.AddHours(1);
            appointment.End = appointment.Start.AddMinutes(30);
            appointment.IsReminderSet = true;
            appointment.ReminderMinutesBeforeStart = 15;
            appointment.RequiredAttendees.Add(new Attendee(userEmail));
            appointment.Save(new FolderId(WellKnownFolderName.Calendar, userEmail), SendInvitationsMode.SendToAllAndSaveCopy);

            // Verify that the appointment was created by using the appointment item ID.
            Item item = Item.Bind(service, appointment.Id, new PropertySet(ItemSchema.Subject));
            Console.WriteLine("Meeting created: " + item.Subject + "\n");
            Console.Read();
            DisplayMessage("\t\t\t\tAppointment created. Press Enter to Continue...");
        }

        static void CancelAppointment()
        {
            DisplayMessage("\t\t\t\tCancelling Appointment, please enter message id...");

            DateTime startDate = DateTime.Now;
            DateTime endDate = startDate.AddMonths(2);
            const int num_appts = 10;

            CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());

            CalendarView cView = new CalendarView(startDate, endDate, num_appts);
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End);

            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

            foreach (Appointment item in appointments.Items)
            {
                Console.WriteLine("Subject: " + item.Subject.ToString() + " ");
                Console.WriteLine("Start: " + item.Start.ToString() + " ");
                Console.WriteLine("End: " + item.End.ToString());
                Console.WriteLine("Message ID: " + item.Id.ToString());
                Console.WriteLine();
            }

            Console.WriteLine(); Console.WriteLine();

            if (appointments.Count() != 0) //check available appointments/meetings 
            {
                Console.WriteLine("Enter the id to cancel meeting: ");
                string id = Console.ReadLine();

                try
                {
                    Appointment meeting = Appointment.Bind(service, id, new PropertySet());

                    CancelMeetingMessage cancelMessage = meeting.CreateCancelMeetingMessage();
                    cancelMessage.Body = new MessageBody("The meeting has been cancelled due to unforeseen circumstances.");
                    cancelMessage.IsReadReceiptRequested = true;
                    cancelMessage.SendAndSaveCopy();
                    Console.WriteLine();
                    DisplayMessage("\t\t\t\tAppointment cancelled. Press Enter to Continue...");
                }
                catch (Exception e)
                {
                    DisplayMessage("\t\t\t\t\t\tInvalid message id");
                }
            }

            else
                DisplayMessage("\t\t\t\t\tNo Appointment found!");
            
        }

        static void FindAllAppointment()
        {
            DisplayMessage("\t\t\t\t\tFinding Appointments, please wait...");

            FindItemsResults<Appointment> findAppointments = service.FindAppointments(WellKnownFolderName.Calendar, new CalendarView(DateTime.Now, DateTime.Now.AddMonths(2)));

            DisplayMessage(String.Format("\t\t\t\tFound {0} appointments. Press Enter to Continue...", findAppointments.Count()));

            foreach (Appointment item in findAppointments)
            {
                Console.WriteLine("Subject: " + item.Subject);
                Console.WriteLine("Start: " + item.Start);
                Console.WriteLine("Duration: " + item.Duration);
                Console.WriteLine("Location: " + item.Location);
                Console.WriteLine("Time Left: " + (item.Start - DateTime.Now));
                Console.WriteLine("---------------------------");
            }
        }

        static void CheckAvailability()
        {
            DisplayMessage("\t\t\t\tCheck Availability, please wait...");
            List<AttendeeInfo> attendees = new List<AttendeeInfo>();
            attendees.Add(new AttendeeInfo(userEmail));

            GetUserAvailabilityResults results = service.GetUserAvailability(attendees,
                new TimeWindow(DateTime.Now, DateTime.Now.AddHours(24)), AvailabilityData.FreeBusy);

            AttendeeAvailability myAvailablity = results.AttendeesAvailability.FirstOrDefault();
            
            if (myAvailablity != null)
            {
                CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
                CalendarView cView = new CalendarView(DateTime.Now, DateTime.Now.AddHours(24));
                FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

                //DisplayMessage(String.Format("\t\t\tYou have {0} appointments/meetings in the next 24 hours. Please Enter to continue...", appointments.Count()));
                DisplayMessage(String.Format("\t\t\tYou have {0} appointments/meetings in the next 24 hours. Please Enter to continue...", myAvailablity.CalendarEvents.Count()));

                foreach (Appointment item in appointments)
                {
                    Console.WriteLine("Subject: " + item.Subject);
                    Console.WriteLine("Start: " + item.Start);
                    Console.WriteLine("Duration: " + item.Duration);
                    Console.WriteLine("Location: " + item.Location);
                    Console.WriteLine("Time Left: " + (item.Start - DateTime.Now));
                    Console.WriteLine("---------------------------");
                }
            }
        }

    }
}
