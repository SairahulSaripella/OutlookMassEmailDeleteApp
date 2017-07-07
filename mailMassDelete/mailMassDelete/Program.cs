using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

// THIS CONSOLE APP ONLY WORKS IF YOU HAVE OUTLOOK ON YOUR DESKTOP AND ARE LOGGED IN


namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            // welcome message
            System.Console.WriteLine("Welcome to the Outlook Mass Mail Delete Applciation \n \n");
            System.Console.WriteLine("Save time clicking delete by specifying a common email address or subject line \n");
            System.Console.WriteLine("======================================================================================== \n \n");

            // this runs the main loop which after each user choice runs itself again through recursion
            run();


        }


        // This function's purpose is to display the menu of what the user can do and how to say a command
        // it then figures out which option the user selected by looking at the number in the first token
        // then it calls the proper function based on whether the program is looking for keyword by subject or by sender email
        public static void run()
        {
            System.Console.WriteLine("1. Delete emails by Subject Keyword \n");
            System.Console.WriteLine("2. Delete emails by Email Keyword \n" +
                                        "PRESS exit to leave \n");


            System.Console.WriteLine("Enter your choice '1 or 2' followed by the subject line keyword or email address you want to start deleting from   \n \n");
            string answer = System.Console.ReadLine();


            // only word to exit out of program is exit
            if (!(answer.Contains("exit")))
            {
                string[] tokens = answer.Split(null as string[], StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length < 2)
                {
                    System.Console.WriteLine("Invalid Command. \n");
                    run();
                    return;
                }
                else if (tokens[0] == "1")
                {
                    massDeleteSubject(tokens[1]);

                    return;
                }
                else if (tokens[0] == "2")
                {
                    massDeleteAddress(tokens[1]);

                    return;
                }
                else
                {
                    System.Console.WriteLine("Invalid Command. \n");
                    run();
                    return;
                }

            }
            else
            {
                exitApp();
            }


        }


        public static void massDeleteSubject(string keyword)
        {

            try
            {
                int count = 0;


                // Connecting to the Mail API namespace lets us access the local outlook endpoint
                Application myApp = new Application();
                NameSpace mapiNameSpace = myApp.GetNamespace("MAPI");
                MAPIFolder myInbox = mapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                // put the length of inbox in variable so as to not mess with the actual myInbox object
                int total = myInbox.Items.Count;
                if (total > 0)
                {

                    // to start from end of mail list (most recent)
                    for (; (total > 0); total = total - 1)
                    {

                        string sender = myInbox.Items[total].Subject;

                        if (sender.Contains(keyword))
                        {
                            MailItem m = myInbox.Items[total];
                            m.Delete();

                            count++;

                        }


                    }

                }

                System.Console.WriteLine("\nTotal mail deleted: " + count + Environment.NewLine);

            }
            catch (System.Exception e)
            {
                System.Console.WriteLine("Error accessing mail occurred. Message: " + e.Message);

            }

            // after deleting emails, user can do the same with another keyword or exit by command
            run();

        }

        // same function as above to access outlook through interops and delete emails
        public static void massDeleteAddress(string keyword)
        {

            try
            {
                int count = 0;

                Application myApp = new Application();
                NameSpace mapiNameSpace = myApp.GetNamespace("MAPI");
                MAPIFolder myInbox = mapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                int total = myInbox.Items.Count;
                if (total > 0)
                {

                    // to start from end of mail list (most recent)
                    for (; (total > 0); total = total - 1)
                    {

                        string email = myInbox.Items[total].SenderEmailAddress;

                        if (email.Contains(keyword))
                        {
                            MailItem m = myInbox.Items[total];
                            m.Delete();

                            count++;

                        }


                    }

                }

                System.Console.WriteLine("Total mail deleted: " + count + Environment.NewLine);

            }   // if there is an error then send error message
            catch (System.Exception e)
            {
                System.Console.WriteLine("Error accessing mail occurred. Message: " + e.Message);

            }

            run();

        }

        public static void exitApp()
        {
            Environment.Exit(0);
        }


    }
}
