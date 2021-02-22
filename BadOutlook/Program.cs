using System;

namespace BadOutlook
{
    class Program
    {
        static void Main(string[] args)
        {
            var mails = OutlookEmails.ReadMailItems();
            int i = 1;

            foreach (var mail in mails)
            {
                Console.WriteLine("Mail No " + i);
                Console.WriteLine("Mail Recieved from " + mail.EmailFrom);
                Console.WriteLine("Mail Subject " + mail.EmailSubject);
                Console.WriteLine("Mail Body " + mail.EmailBody);
                Console.WriteLine("");

                i += 1;

            }

            Console.ReadKey();
        }
    }
}
