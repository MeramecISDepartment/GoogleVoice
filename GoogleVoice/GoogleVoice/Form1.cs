//Google Calendar related namespaces
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

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

        //This will be used for the alarm.
        static string now;

        private void Form1_Load(object sender, EventArgs e)
        {
            //if (credPath != System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal))
            //{
            //These web pages helped me to learn to delete files.
            //https://msdn.microsoft.com/en-us/library/cc148994(v=vs.100).aspx
            //http://stackoverflow.com/questions/8861854/read-a-file-from-an-unknown-location
            //System.IO.File.Delete(@"C:\Users\Tyler Barton\Documents\.credentials\calendar-dotnet-quickstart.json\Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            System.IO.File.Delete(@"Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            //}
            Choices voiceCommands = new Choices();
            voiceCommands.Add(new string[] { "Add an event to my Google Calendar.", "View the next 10 upcoming events.", /*"Set my pre-planned alarm clock."*/ });
            GrammarBuilder grammarBuilder = new GrammarBuilder();
            grammarBuilder.Append(voiceCommands);
            Grammar grammar = new Grammar(grammarBuilder);

            speechRecEngine.LoadGrammarAsync(grammar);
            speechRecEngine.SetInputToDefaultAudioDevice();
            speechRecEngine.SpeechRecognized += speechRecEngine_SpeechRecognized;
            synthesizer.SpeakAsync("Hi, please long into your Google Calendar.");

            //This code will log you into Google Calendar before continuing forward.
            //modified code from the .Net Quickstart
            string[] Scopes = { CalendarService.Scope.CalendarReadonly };
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

            speechRecEngine.RecognizeAsync(RecognizeMode.Multiple);
        }


        //Speech recognition related method
        void speechRecEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text)
            {
                case "Add an event to my Google Calendar":
                    break;
                case "View the next 10 upcoming events.":
                    // Define parameters of request.
                    EventsResource.ListRequest request = service.Events.List("primary");
                    request.TimeMin = DateTime.Now;
                    request.ShowDeleted = false;
                    request.SingleEvents = true;
                    request.MaxResults = 10;
                    request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

                    // List events.
                    Events events = request.Execute();
                    richTextBox1.Text = "Upcoming events:";

                    synthesizer.SpeakAsync("Your upcoming events are:");
                    if (events.Items != null && events.Items.Count > 0)
                    {
                        foreach (var eventItem in events.Items)
                        {
                            string when = eventItem.Start.DateTime.ToString();
                            if (String.IsNullOrEmpty(when))
                            {
                                when = eventItem.Start.Date;

                                if (DateTime.Now.CompareTo(when) == 0)
                                {
                                    now = when;
                                }
                            }
                            synthesizer.SpeakAsync(eventItem.Summary + " on " + when + ".");
                            richTextBox1.Text += "\n" + eventItem.Summary + " (" + when + ")";
                        }
                    }
                    else
                    {
                        synthesizer.SpeakAsync("You have no upcoming events.");
                        richTextBox1.Text = "No upcoming events found.";
                    }
                    //Console.Read();
                    break;

                    /*case "Set my pre-planned alarm clock." :

                        if (richTextBox1.Text == "")
                        {
                            synthesizer.SpeakAsync("You must say 'view the next 10 upcoming events' first.");
                            break;
                        }

                        //I referenced this for the alarm: 
                        //apps.topcoder.com/forums/?module=Thread&threadID=670237
                        synthesizer.SpeakAsync("Your next alarm will be at " + now);
                        synthesizer.SpeakAsync("You will hear the following sound:");

                        //I got this code from here:
                        //http://stackoverflow.com/questions/3502311/how-to-play-a-sound-in-c-net
                        System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"C:\Users\Tyler Barton\Desktop\GoogleVoice\GoogleVoice\GetUp.wav");
                        player.Play();

                        if (DateTime.Now.CompareTo(now) >= 0)
                        {
                            //I got this code from here:
                            //http://stackoverflow.com/questions/3502311/how-to-play-a-sound-in-c-net
                            player.Play();
                        }
                        break;*/
            }
        }
    }
}
