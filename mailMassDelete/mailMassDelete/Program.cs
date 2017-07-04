using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
 *          This Program can only run if you have microsoft outlook installed as desktop app
 *          on your system and you are logged in. This is only meant to make deleting large amounts
 *          of useless email quickly.
 * 
 * 
 * */


namespace mailMassDelete
{
    class Program
    {
        static void Main(string[] args)
        {
            // welcome message
            System.Console.WriteLine("Welcome to the Outlook Mass Mail Delete Applciation \n \n");
            System.Console.WriteLine("Save time clicking delete by specifying a common email address or subject line and how many iterations you want deleted \n");
            System.Console.WriteLine("======================================================================================================================== \n \n");





            // run function uses recrusion to keep calling itself if the user 
            // enters any invalid answer or the delete function runs to allow the user to keep entering
            // keywords to delete
            run();


        }

        public static void run()
        {
            System.Console.WriteLine("1. Delete by Subject \n" +
                                     "2. Delete by sender email");
            System.Console.WriteLine("Enter your choice '1 or 2' followed by the subject line keyword or email address you want to start deleting from   \n \n");
            string read = System.Console.ReadLine();
            if (!(read.Contains("exit")))
            {

                string[] tokens = read.Split(null as string[], StringSplitOptions.RemoveEmptyEntries);

                if (tokens[0] == "1")
                {
                    massDeleteBySubject(tokens[1]);

                }
                else if (tokens[0] == "2")
                {
                    massDeleteByEmail(tokens[1]);
                }
                else
                {
                    System.Console.WriteLine("Invalid choice \n");

                    run();
                }



            }
            else
            {
                exitApp();
            }
        }

        public static void massDeleteBySubject(string keyword)
        {
            // this function uses the keyword to iterate through your mail and delete matches, then it runs the menu again
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

                        string sender = myInbox.Items[total].Subject;

                        if (sender.Contains(keyword))
                        {
                            MailItem m = myInbox.Items[total];
                            m.Delete();

                            count++;

                        }


                    }

                }

                System.Console.WriteLine("Total mail deleted: " + count);

            }
            catch (System.Exception e)
            {
                System.Console.WriteLine("Error accessing mail occurred. Message: " + e.Message);

            }


            run();

        }


        public static void massDeleteByEmail(string keyword)
        {
            if (!(keyword.Contains("@")))
            {
                System.Console.WriteLine("Invalid email. Action failed.");
                return;
            }
           
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

                        string sender = myInbox.Items[total].SenderEmailAdress;

                        if (sender.Contains(keyword))
                        {
                            MailItem m = myInbox.Items[total];
                            m.Delete();

                            count++;

                        }


                    }

                }

                System.Console.WriteLine("Total mail deleted: " + count);

            }
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
