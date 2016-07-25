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
    class deleteEvent
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

        public void deleteAnEvent()
        {
            //I referenced Google Developer and Linda Lawton for this:
            //https://github.com/LindaLawton/Google-Dotnet-Samples/blob/master/Google-Calendar/Google-Calendar-Api-dotnet/Google-Clndr-Api-dotnet/Google-Clndr-Api-dotnet/DaimtoEventHelper.cs
            // Initialize Calendar service with valid OAuth credentials
            //Calendar service = new Calendar.Builder(httpTransport, jsonFactory, credentials)
            //    .setApplicationName("applicationName").build();

            // Delete an event
            //service.Events.Delete("primary", "Google I/O 2015").deleteAnEvent.execute;
            //service.events().delete('primary', "eventId").execute();
            //I will reference this kind soul in the future:
            //https://groups.google.com/forum/#!topic/google-calendar-api/gZO18rvvqHM

            //try
            //{
             //   return service.Events.Delete('primary', 'Google I/O 2015').Execute();
            //}
            //catch (Exception ex)
            //{
                //Console.WriteLine(ex.Message);
                //return null;
            //}
        }
    }
}