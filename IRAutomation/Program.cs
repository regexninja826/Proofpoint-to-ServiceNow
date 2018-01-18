using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Threading;
using Newtonsoft.Json;

namespace IRAutomation
{
    class Program
    {
        protected static string snpass;
        protected static string snuser;
        protected static string mailuser;
        protected static string mailpass;
        private static string proxy;
        
        static void Main(string[] args)
        {
            snpass = args[0];
            snuser = args[1];
            mailuser = args[2];
            mailpass = args[3];
            if (args.Length==5) {
                proxy = args[4];
            }

            bool p = false;
            if (proxy == "proxy")
            {

                p = true;
            }


            while (true)
            {

                //this could use another method for calling every 5 minutes
                //suggestions are welcome
                

                int minute = int.Parse(DateTime.Now.Minute.ToString());
                if (minute % 5 == 0)
                {
                    Console.WriteLine(minute);
                    //here for trouble shooting




                   ;


                    EventGenerator eventgen = new EventGenerator(mailuser, mailpass, snuser, snpass);

                    eventgen.generateevents(p);
                    Thread.Sleep(1000 * 60 * 2);
                   

                }
                               
                
            }


            /*
             * you need this if you are behind the proxy.... along with cygwin
             *socat TCP4-LISTEN:2022,reuseaddr,fork \PROXY:yourProxy:outlook.office365.com,proxyport=yourproxyport
             * 
             * yes there ae better ways to work with a proxy. feel free to incorporate or share your own methods
             */





        }
    }
}
;