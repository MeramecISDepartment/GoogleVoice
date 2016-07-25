//Google Calendar related namespaces
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
//using Google.Apis;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading; //This is used in the voice to text application
using System.Windows.Forms;

namespace GoogleVoice
{

    class addEvent
    {
    
        public CalendarService service;
            
        public void login()
        {
            string[] Scopes = { CalendarService.Scope.Calendar };
            string ApplicationName = "Google Calendar API .NET Quickstart";

            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/calendar-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //synthesizer.SpeakAsync("Credential file saved to: " + credPath);
                //richTextBox1.Text = "Credential file saved to: " + credPath;
            }

            // Create Google Calendar API service.
            service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        public void addNewEvent()
        {
            

            // Refer to the .NET quickstart on how to setup the environment:
            // https://developers.google.com/google-apps/calendar/quickstart/dotnet
            // Change the scope to CalendarService.Scope.Calendar and delete any stored
            // credentials.
            ;
            Event newEvent = new Event()
            {
                Summary = "Google I/O 2015",
                //prompt for summary, verify and store
                Location = "800 Howard St., San Francisco, CA 94103",
                //prompt for address
                //prompt for city
                //prompt for state
                //prompt for zipcode
                //make a string that says address + ", " + city + ", " + state + " " + zipcode
                //have it so that if any state the initials will shop up
                //verify with each one before saving.

                Description = "A chance to hear more about Google's developer products.",
                Start = new EventDateTime()
                {
                    DateTime = DateTime.Parse("2015-05-28T09:00:00-07:00"),
                    TimeZone = "America/Los_Angeles",
                },
                End = new EventDateTime()
                {
                    DateTime = DateTime.Parse("2015-05-28T17:00:00-07:00"),
                    //prompt year and save it
                    //prompt month and save it
                    //prompt day and save it
                    //prompt start hour and save it (12/12 instead of 24)
                    //take extra at the end and save it to minutes, o'clock (00)
                    //seconds at the end should be saved (none, 00 seconds)
                    //prompt end time and save it 
                    //take extra at the end and save it to minutes
                    //take extra seconds at the end and save them
                    //put a string together as year + "-" + month + "-" + day + "T" + hour1 + ":" + minutes1 + ":" + seconds "-" + hour2 + ":" + minutes2 + ":" + seconds2

                    TimeZone = "America/Los_Angeles",
                    //Have a list of time zones
                },
                Recurrence = new String[] { "RRULE:FREQ=DAILY;COUNT=2" },
                //How frequent would you like for it to occur store the number
                //once does not repeat the event
                //daily/monthly will store the choice if not either will return an error 
                //save a string as "RRULE:FREQ=" + interval + ";COUNT=" + frequency

                Attendees = new EventAttendee[]
                {
                                 new EventAttendee() { Email = "lpage@example.com" },
                                 //prompt for an email
                                 //verify that it is in the proper format *****@***.*** (org, net, com, edu, etc.)
                                 //verify that it exists and save/
                                 //prompt for another
                                 //etc 
                                 //maybe save to a collection or a list
                                 new EventAttendee() { Email = "sbrin@example.com" },
                },
                Reminders = new Event.RemindersData()
                {
                    UseDefault = false,
                    Overrides = new EventReminder[]
                    {
                                  new EventReminder() { Method = "email", Minutes = 24 * 60 },
                                  //prompt method (email, sms, other)
                                  //prompt minutes
                                  //prompt for (what 60 is)
                                  new EventReminder() { Method = "sms", Minutes = 10 },
                    }
                }
            };

            String calendarId = "primary";
            EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
            Event createdEvent = request.Execute();
        }
    }
}
