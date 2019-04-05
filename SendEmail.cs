using System;
using System.Collections.Generic;

namespace BirthdayPrg
{
    internal class SendEmail
    {
        private dynamic Notes = null;        
        private dynamic NotesDocument = null;

        private string userName = null;
        private string MailDbName = null;

        //testing img db email
        private dynamic body = null;

        private dynamic header = null;
        private dynamic child = null;
        private dynamic stream = null;
        private string fileFormat = "";
        private dynamic rtItemA = null;
        private dynamic rtItemB = null;
        //

        public void StartNotes()
        {
            //Start a session to Notes

            Notes = Activator.CreateInstance(Type.GetTypeFromProgID("Notes.NotesSession"));

            userName = Notes.userName;

            MailDbName = userName.Substring(3, 1) + userName.Substring(userName.Length - ((userName.Length - (userName.IndexOf(" ", 0) + 1)))) + ".nsf";
        }

        public void SendNotesMail(string bdayName, string bdayEmail)
        {
            dynamic testdb = Notes.CurrentDatabase;

            NotesDocument = testdb.CreateDocument();
            object oItemValue = null;
            string strImg = "C:\\Users\\michael.gonzalez\\source\\repos\\BirthdayPrg\\BirthdayPrg\\Resources\\happy-birthday.bmp";
            string Recipient = bdayEmail;
            NotesDocument.replaceItemValue("Form", "Memo");
            NotesDocument.replaceItemValue("Principal", "hr@baccredomatic.us");
            NotesDocument.replaceItemValue("Subject", "Happy Birthday!");
            NotesDocument.replaceItemValue("SendTo", bdayEmail);
            NotesDocument.replaceItemValue("From", "hr@baccredomatic.us");
            oItemValue = NotesDocument.GetItemValue("SendTo");
            dynamic notesRTItem = NotesDocument.createRichTextItem("Body");

            stream = Notes.CreateStream;
            stream.Open(strImg);
            body = NotesDocument.CreateMIMEEntity("DummyRichText");
            header = body.CreateHeader("Content-Type");
            header.SetHeaderVal("multipart/mixed");
            child = body.CreateChildEntity();
            fileFormat = "image/bmp";
            child.SetContentFromBytes(stream, fileFormat, 1730);
            stream.Close();
            NotesDocument.Save(false, false);
            rtItemA = NotesDocument.GetFirstItem("Body");
            rtItemB = NotesDocument.GetFirstItem("DummyRichText");
            rtItemA.AppendRTItem(rtItemB);
            rtItemB.Remove();
            NotesDocument.Save(false, false);

            notesRTItem.AppendText("\nHappy Birthday, " + bdayName + "!\n\n");
            notesRTItem.AppendText("We appreciate you here and thank you for your continued ");
            notesRTItem.AppendText("support and work.\n");
            NotesDocument.Save(true, false);
            NotesDocument.Send(false, ref oItemValue);
        }

        public void SendNotesMailWeek(List<string[]> matches)
        {
            dynamic testdb = Notes.CurrentDatabase;

            NotesDocument = testdb.CreateDocument();
            object oItemValue = null;
            string strImg = "C:\\Users\\michael.gonzalez\\source\\repos\\BirthdayPrg\\BirthdayPrg\\Resources\\happy-birthday-logo.bmp";
            string Recipient = "";
            NotesDocument.replaceItemValue("Form", "Memo");
            NotesDocument.replaceItemValue("Principal", "hr@baccredomatic.us");
            NotesDocument.replaceItemValue("Subject", "Happy Birthday!");
            NotesDocument.replaceItemValue("SendTo", "hr@baccredomatic.us");
            NotesDocument.replaceItemValue("From", "hr@baccredomatic.us");
            oItemValue = NotesDocument.GetItemValue("SendTo");
            dynamic notesRTItem = NotesDocument.createRichTextItem("Body");

            stream = Notes.CreateStream;
            stream.Open(strImg);
            body = NotesDocument.CreateMIMEEntity("DummyRichText");
            header = body.CreateHeader("Content-Type");
            header.SetHeaderVal("multipart/mixed");
            child = body.CreateChildEntity();
            fileFormat = "image/bmp";
            child.SetContentFromBytes(stream, fileFormat, 1730);
            stream.Close();
            NotesDocument.Save(false, false);
            rtItemA = NotesDocument.GetFirstItem("Body");
            rtItemB = NotesDocument.GetFirstItem("DummyRichText");
            rtItemA.AppendRTItem(rtItemB);
            rtItemB.Remove();
            NotesDocument.Save(false, false);

            notesRTItem.AppendText("\nThis week we have the following birthdays, \n\n");
            foreach (var item in (matches))
            {
                notesRTItem.AppendText(item[0].ToString() + " " + item[1].ToString() + ", on " + item[3].ToString() + " / " + item[4].ToString());
                notesRTItem.AppendText("\n");
            }
            notesRTItem.AppendText("\n\nMake sure to wish them a happy birthday!\n");
            NotesDocument.Save(true, false);
            NotesDocument.Send(false, ref oItemValue);
        }

        public void CloseNotes()
        {
            // UIdoc = WorkSpace.currentdocument;

            // UIdoc.Close();
        }
    }
}