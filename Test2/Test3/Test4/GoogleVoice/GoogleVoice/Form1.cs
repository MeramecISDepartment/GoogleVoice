using System;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading; //This is used in the voice to text application
using System.Threading.Tasks;
using System.Windows.Forms;


//Speech recognition related namespaces
using System.Speech; //This is used in the voice to text application
using System.Speech.Recognition; //This is used in the voice to text application
using System.Speech.Synthesis; //This is used in the voice to text application

//Audio namespaces

namespace GoogleVoice
{

    public partial class Form1 : Form
    {

        public Thread RecThread; //New
        public Boolean RecognizerState = true; //New

        public string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);

        public Form1()
        {
            InitializeComponent();
        }

        //string credPath = System.Environment.GetFolderPath(
        //                    System.Environment.SpecialFolder.Personal);

        //Speech recognition related items
        SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();

        private void Form1_Load(object sender, EventArgs e)
        {
            //if (credPath != System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal))
            //{
            //These web pages helped me to learn to delete files.
            //https://msdn.microsoft.com/en-us/library/cc148994(v=vs.100).aspx
            //http://stackoverflow.com/questions/8861854/read-a-file-from-an-unknown-location
            //System.IO.File.Delete(@"C:\Users\Tyler Barton\Documents\.credentials\calendar-dotnet-quickstart.json\Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            credPath = Path.Combine(credPath, ".credentials/calendar-dotnet-quickstart.json");
            //System.IO.File.Delete(@credPath);
            //System.IO.File.Delete(@".credentials/calendar-dotnet-quickstart.json");
            //System.IO.File.Delete(@"calendar-dotnet-quickstart.json");
            //string extension = Path.GetExtension(Directory.GetFiles(NETWORK_FOLDER_PATH, "calendar-dotnet-quickstart.json")[0]);

            /*
             * I think that this might be the solution for deleting a file that's at a different location on each computer.
             * http://stackoverflow.com/questions/1240373/how-do-i-get-the-current-username-in-net-using-c
             * by jaun
             */
            //string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string userName1 = System.Environment.UserName;
            //string credentials = "C:/Users/" + userName + "/Documents/.credentials/calendar-dotnet-quickstart.json/Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";
            System.IO.File.Delete(@"C:/Users/" + userName1 + "/Documents/.credentials/calendar-dotnet-quickstart.json/Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");

            //System.IO.File.Delete(@"Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            //}

            Choices voiceCommands = new Choices();
            voiceCommands.Add(new string[] { "Add an event to my Google Calendar.", "Delete an event from my Google Calendar."});
            GrammarBuilder grammarBuilder = new GrammarBuilder();
            //grammarBuilder.AppendDictation(); //New
            grammarBuilder.Append(voiceCommands);
            Grammar grammar = new Grammar(grammarBuilder);

            recognizer.LoadGrammarAsync(grammar);
            recognizer.SetInputToDefaultAudioDevice();
            //recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized); //New
            recognizer.SpeechRecognized += recognizer_SpeechRecognized;

            recognizer.RecognizeAsync(RecognizeMode.Multiple);

            //RecognizerState = true;
            //RecThread = new Thread(new ThreadStart(RecThreadFunction)); //?
            //RecThread.Start();

            synthesizer.SpeakAsync("What would you like to do with Google Calendar.");
        }


        //Speech recognition related method
        void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (!RecognizerState)
                return;

            this.Invoke((MethodInvoker)delegate
            {
                // richTextBox1.Text += ("" + e.Result.Text.ToLower());

                switch (e.Result.Text)
                {
                    case "Add an event to my Google Calendar.":
                        synthesizer.SpeakAsync("Please give your credentials first.");
                        addEvent newEvent = new addEvent();
                        newEvent.login();
                        synthesizer.SpeakAsync("Your credentials have been saved to " + credPath + ".");
                        synthesizer.SpeakAsync("You have logged in.  What would you like to name your event?");
                        synthesizer.SpeakAsync("Did you say ...");


                        synthesizer.SpeakAsync("What description would you like to give your event?");
                        synthesizer.SpeakAsync("Did you say ...");

                        //synthesizer.SpeakAsync("When does your event begin?");

                        //word numbers will be converted to and saved as numbers using swithc cases //same with states to abbreviations to save
                        //all inputs should be verified back to the user in words until the user says that they are okay
                        //if there is no input or an error in input, the user should be prompted to fill the field out properly

                        synthesizer.SpeakAsync("During what year does your event start?"); //perhaps look at whether the event month has already passed and if so go to the next year
                        synthesizer.SpeakAsync("Did you say ...");
                        synthesizer.SpeakAsync("During what month does your event start?");
                        synthesizer.SpeakAsync("Did you say ...");
                        synthesizer.SpeakAsync("During what day does your event start?");
                        synthesizer.SpeakAsync("Did you say ...");
                        synthesizer.SpeakAsync("During what time does your event start and is it AM or PM??"); //Note, since this is military time 12:00AM - 11:59AM is 1:00 to 12:59 and 12:00PM to 11:59PM is 13:00 to 24:00
                        synthesizer.SpeakAsync("Did you say ...");

                        synthesizer.SpeakAsync("During what year does your event end?");  //perhaps look at whether the event month has already passed or is before the start month and go to the next year
                        synthesizer.SpeakAsync("Did you say ...");
                        synthesizer.SpeakAsync("During what month does your event end");
                        synthesizer.SpeakAsync("Did you say ...");
                        synthesizer.SpeakAsync("During what day does your event end?");
                        synthesizer.SpeakAsync("Did you say ...");
                        synthesizer.SpeakAsync("During what time does your event end and is it AM or PM??");
                        synthesizer.SpeakAsync("Did you say ...");
                        //synthesizer.SpeakAsync("When does your event end?");


                        synthesizer.SpeakAsync("In what time zone are you in?"); //perhaps get it off the computer clock
                        synthesizer.SpeakAsync("Did you say ...");

                        synthesizer.SpeakAsync("How often would you like for this event to repeat (never, daily, bi-daily, tri-daily, weekly, bi-weekly, tri-weekly, monthly, yearly)?");
                        synthesizer.SpeakAsync("Did you say ...");

                        synthesizer.SpeakAsync("Would you to add any attendees (yes or no)?");
                        synthesizer.SpeakAsync("Did you say ...");

                        synthesizer.SpeakAsync("Would is the attendee's email address?");
                        synthesizer.SpeakAsync("Did you say ...");

                        synthesizer.SpeakAsync("You you like to add another attendee (yes or no)?");
                        synthesizer.SpeakAsync("Did you say ...");

                        synthesizer.SpeakAsync("Would you like to add reminders (yes or no)?");
                        synthesizer.SpeakAsync("Did you say ...");

                        synthesizer.SpeakAsync("What type of reminder would you like to give (text or email)?"); //Note: When saved, text = sms
                        synthesizer.SpeakAsync("Did you say ...");

                        synthesizer.SpeakAsync("How much time before the event would you like the reminder (days, hours, minutes, seconds)?"); //Note: 24 * 60 is 1 day
                        //synthesizer.SpeakAsync("How many minutes, hours, or days before the event would you like your reminder to be?");
                        synthesizer.SpeakAsync("Did you say ...");

                        //synthesizer.SpeakAsync("");
                        newEvent.addNewEvent();
                        synthesizer.SpeakAsync("Your new event has been added."); //Give a message with the different information.
                        break;

                    case "Delete an event from my Google Calendar.":
                        deleteEvent delEvent = new deleteEvent();
                        delEvent.deleteAnEvent();
                        break;
                }
            });
        }

        public void RecThreadFunction()
        {
            while (true)
            {
                try
                {
                    recognizer.Recognize();
                }
                catch
                {
                   // synthesizer.SpeakAsync("I did not understand you.");
                }
            }
        }
    }
}
