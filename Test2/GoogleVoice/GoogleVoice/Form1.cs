//Google Calendar related namespaces
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
//using Google.Apis;

using System;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

//Speech recognition related namespaces
using System.Speech;
using System.Speech.Recognition;
using System.Speech.Synthesis;

//Audio namespaces


namespace GoogleVoice
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //string credPath = System.Environment.GetFolderPath(
        //                    System.Environment.SpecialFolder.Personal);

        //Speech recognition related items
        SpeechRecognitionEngine speechRecEngine = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();

        private void Form1_Load(object sender, EventArgs e)
        {
            //if (credPath != System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal))
            //{
            //These web pages helped me to learn to delete files.
            //https://msdn.microsoft.com/en-us/library/cc148994(v=vs.100).aspx
            //http://stackoverflow.com/questions/8861854/read-a-file-from-an-unknown-location
            System.IO.File.Delete(@"C:\Users\Tyler Barton\Documents\.credentials\calendar-dotnet-quickstart.json\Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            //System.IO.File.Delete(@"Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            //}
            Choices voiceCommands = new Choices();
            voiceCommands.Add(new string[] { "Add an event to my Google Calendar."});
            GrammarBuilder grammarBuilder = new GrammarBuilder();
            grammarBuilder.Append(voiceCommands);
            Grammar grammar = new Grammar(grammarBuilder);

            speechRecEngine.LoadGrammarAsync(grammar);
            speechRecEngine.SetInputToDefaultAudioDevice();
            speechRecEngine.SpeechRecognized += speechRecEngine_SpeechRecognized;
            synthesizer.SpeakAsync("What would you like to do with Google Calendar.");
            speechRecEngine.RecognizeAsync(RecognizeMode.Multiple);
        }


        //Speech recognition related method
        void speechRecEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text)
            {
                case "Add an event to my Google Calendar.":
                    //This code will log you into Google Calendar before continuing forward.
                    //modified code from the .Net Quickstart
                    string[] Scopes = { CalendarService.Scope.Calendar };
                    string ApplicationName = "Google Calendar API .NET Quickstart";

                    UserCredential credential;

                    synthesizer.SpeakAsync("Please give your credentials first.");
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
                        synthesizer.SpeakAsync("You have logged in.  What would you like to do next?");
                    }

                    // Create Google Calendar API service.
                    var service = new CalendarService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });


                    // Refer to the .NET quickstart on how to setup the environment:
                    // https://developers.google.com/google-apps/calendar/quickstart/dotnet
                    // Change the scope to CalendarService.Scope.Calendar and delete any stored
                    // credentials.

                    Event newEvent = new Event()
                    {
                        Summary = "Google I/O 2015",
                        Location = "800 Howard St., San Francisco, CA 94103",
                        Description = "A chance to hear more about Google's developer products.",
                        Start = new EventDateTime()
                        {
                            DateTime = DateTime.Parse("2015-05-28T09:00:00-07:00"),
                            TimeZone = "America/Los_Angeles",
                        },
                        End = new EventDateTime()
                        {
                            DateTime = DateTime.Parse("2015-05-28T17:00:00-07:00"),
                            TimeZone = "America/Los_Angeles",
                        },
                        Recurrence = new String[] { "RRULE:FREQ=DAILY;COUNT=2" },
                        Attendees = new EventAttendee[] {
        new EventAttendee() { Email = "lpage@example.com" },
        new EventAttendee() { Email = "sbrin@example.com" },
    },
                        Reminders = new Event.RemindersData()
                        {
                            UseDefault = false,
                            Overrides = new EventReminder[] {
            new EventReminder() { Method = "email", Minutes = 24 * 60 },
            new EventReminder() { Method = "sms", Minutes = 10 },
        }
                        }
                    };

                    String calendarId = "primary";
                    EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
                    Event createdEvent = request.Execute();
                    richTextBox1.Text = "Event created: " + createdEvent.HtmlLink;
                    richTextBox1.Text ="The event, " + newEvent.Summary + " has been added to your Google Calendar.  " +
                         "It is held at " + newEvent.Location + " from " + newEvent.Start.DateTime + " until " + newEvent.End.DateTime +
                         "The event is described as " + newEvent.Description + "." /*f+ "The number of recourances of this event are " /*+ newEvent.Recurrence.Count + 
                         "The number of attendees are " + newEvent.Attendees.Count + "The next reminders will be " + newEvent.Reminders.Overrides*/;
                    synthesizer.SpeakAsync("The event, " + newEvent.Summary + " has been added to your Google Calendar.  " +
                         "It is held at " + newEvent.Location + " from " + newEvent.Start.DateTime + " until " + newEvent.End.DateTime + 
                         "The event is described as " + newEvent.Description + "." /*+ "The number of recourances of this event are " /*+ newEvent.Recurrence.Count + 
                         "The number of attendees are " + newEvent.Attendees.Count + "The next reminders will be " + newEvent.Reminders.Overrides*/);
                    break;
            }
        }
    }
}
