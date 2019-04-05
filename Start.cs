using System;
using System.Collections.Generic;

namespace BirthdayPrg
{
    internal class Start
    {
        public static void Main()
        {
            List<string[]> matches = Read_From_Excel.getExcelFile();

            string fullName = "";
            string email = "";
            try
            {
                if (matches.Count > 0)
                {
                    //start Notes session to send emails
                    SendEmail NotesSession = new SendEmail();
                    NotesSession.StartNotes();

                    //for each person, send them an email
                    foreach (var item in (matches))
                    {
                        fullName = item[0].ToString() + " " + item[1].ToString();
                        email = item[2].ToString();
                        NotesSession.SendNotesMail(fullName, email);
                    }
                    //close up each new msg
                    foreach (var item in (matches))
                    {
                        NotesSession.CloseNotes();
                    }
                }

                //check for day of week to run weekly email. we look for Monday
                string wk = DateTime.Today.DayOfWeek.ToString();

                if (wk.Equals("Monday"))
                {
                    List<string[]> weekMatches = Read_From_Excel.getWeekExcelFile();

                    if (weekMatches.Count > 0)
                    {
                        //start Notes session to send emails
                        SendEmail NotesSessionWeek = new SendEmail();
                        NotesSessionWeek.StartNotes();

                        //for each person, send them an email

                        NotesSessionWeek.SendNotesMailWeek(weekMatches);

                        //close up each new msg
                        //NotesSessionWeek.CloseNotes();
                    }
                }
            } catch(Exception ex)
            {
                Console.WriteLine("It looks like an error occurred. Please make sure you are logged into Lotus Notes and retry.");
                Console.ReadLine();
                Environment.Exit(0);
            }
            //Console.ReadLine();
            Environment.Exit(0);
        }
    }
}