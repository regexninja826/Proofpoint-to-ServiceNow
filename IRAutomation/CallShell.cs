using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.IO;
using System.Collections;
using System.Collections.ObjectModel;

namespace IRAutomation
{
    class CallShell
    {
        public CallShell()
        {

        }

        public Collection<PSObject> execute_posh(string s, string upn)
        {
            RunspaceConfiguration rc = RunspaceConfiguration.Create();
            Runspace runspace = RunspaceFactory.CreateRunspace(rc);
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();

            //TextReader txtr = new StreamReader(@"c:\test\somePSscript.ps1");

            TextReader txtr = new StreamReader(s);
            if (txtr != null)
            {
                string str1 = txtr.ReadToEnd();
                Command myCommand = new Command(str1, true);
                CommandParameter userParam = new CommandParameter("upn", upn);
                myCommand.Parameters.Add(userParam);
                pipeline.Commands.Add(myCommand);
            }

            Collection<PSObject> results = pipeline.Invoke();
            runspace.Close();
            return results;
        }

    }
}
