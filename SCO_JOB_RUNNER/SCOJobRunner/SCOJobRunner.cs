using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data.Services.Client;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.SystemCenter.Orchestrator.Integration.Administration.SCOJobRunner.OrchestratorWebService;


namespace Microsoft.SystemCenter.Orchestrator.Integration.Administration.SCOJobRunner
{
    class SCOJobRunner
    {

        private static Guid _runbookID;
        private static string _runbookPath = string.Empty;
        private static string _webServer = Environment.MachineName;
        private static string _runbookServer = string.Empty;
        private static Dictionary<string, string> _jobParameters = null;
        private static string _port = "81";
        private static bool _useSSL = false;
        private static NetworkCredential _credentials = null;
        private static string _user = string.Empty;
        private static string _domain = string.Empty;
        private static string _password = string.Empty;
        private static bool _verbose = false;
        private static string _serviceRoot = string.Empty;
        private static OrchestratorWebService.OrchestratorContext _context = null;


        static void Main(string[] args)
        {
            if ((args.Length == 0) || (args[0] == "/?") || (args[0].ToLower() == "-help") || (args[0].ToLower() == "/help"))
            {
                PrintUsage();
                return;
            }

            //Parse the arguments
            Arguments CommandLine = new Arguments(args);

            if (CommandLine["Verbose"] != null)
            {
                _verbose = true;
            }

            if (CommandLine["ID"] != null)
            {
                _runbookID = new Guid(CommandLine["ID"]);
            }
            else
            {
                if (CommandLine["RunbookPath"] != null)
                {
                    _runbookPath = CommandLine["RunbookPath"];
                }
                else
                {
                    Console.WriteLine("ID parameter was not provided.");
                    PrintUsage();
                    return;
                }
            }

           

            if (CommandLine["WebServer"] != null)
            {
                _webServer = CommandLine["WebServer"];
            }

            if (CommandLine["Port"] != null)
            { 
                _port = CommandLine["Port"];
            }

            if (CommandLine["UseSSL"] != null)
            {
                _useSSL = true;
            }

            if (CommandLine["RunbookServer"] != null)
            {
                _runbookServer = CommandLine["RunbookServer"];
            }

            if (CommandLine["User"] != null)
            {
                _user = CommandLine["User"];
            }

            if (CommandLine["Domain"] != null)
            {
                _domain = CommandLine["Domain"];
            }

            if (CommandLine["Password"] != null)
            {
                _password = CommandLine["Password"];
            }

            //if any are provided, but not all provided, print usage
            if ((!string.IsNullOrEmpty(_user)) || (!string.IsNullOrEmpty(_domain)) || (!string.IsNullOrEmpty(_password)))
            {
                if ((string.IsNullOrEmpty(_user)) || (string.IsNullOrEmpty(_domain)) || (string.IsNullOrEmpty(_password)))
                {
                    Console.WriteLine("\n\nOne or more of the user, domain and password parameters was empty.\n\n");
                    PrintUsage();
                    return;
                }
                else
                {
                    _credentials = new NetworkCredential(_user, _password, _domain);
                }
            }

            if (CommandLine["GetParameters"] != null)
            {

                Console.WriteLine(GetRunbookParameters());
                return;

            }

            if (CommandLine["Parameters"] != null)
            {

                //TODO: Add code here to verify that the parameters provided match parameters available

                _jobParameters = new Dictionary<string, string>();
                string[] parameterSets = CommandLine["Parameters"].Split(';');
                foreach (string parameterSet in parameterSets)
                {
                    string[] pair = parameterSet.Split('=');
                    _jobParameters.Add(pair[0], pair[1]);
                }
            }

            InvokeRunbook();
        }

        public static void Connect()
        {
            // Path to Orchestrator web service
            string uriRoot = "http";
            if (_useSSL)
            {
                uriRoot = "https";
            }
            _serviceRoot = string.Format(@"{0}://{1}:{2}/Orchestrator2012/Orchestrator.svc", uriRoot, _webServer, _port);

            // Create Orchestrator context
            _context = new OrchestratorWebService.OrchestratorContext(new Uri(_serviceRoot));

            // Set credentials to default or a specific user.
            if (null == _credentials)
            {
                _context.Credentials = System.Net.CredentialCache.DefaultCredentials;
            }
            else
            {
                _context.Credentials = _credentials;
            }


        }

        public static string InvokeRunbook()
        {
            Connect();
            string response = string.Empty;
            string parametersXml = string.Empty;
            if (_verbose)
            {
                Console.WriteLine("\n\nGetting runbook parameters from the web service.");
            }

            try
            {
                if (!string.IsNullOrEmpty(_runbookPath))
                {
                    _runbookID = GetRunbookIdFromPath(_context, _runbookPath);  
                }

                parametersXml = FormatRunbookParameterString(_context, _runbookID, _jobParameters);

            }
            catch (Exception e)
            {
                Console.WriteLine("\n\nError getting runbook parameters from the web service.\n\n" + e);
            }

            try
            {
                if (_verbose)
                {
                    Console.WriteLine("\n\nCreating the new job");
                }

                // Create new job and assign runbook Id and parameters.
                Job job = new Job();
                job.RunbookId = _runbookID;
                if (!string.IsNullOrEmpty(parametersXml))
                {
                    job.Parameters = parametersXml.ToString();
                }
                if (!string.IsNullOrEmpty(_runbookServer))
                {
                    RunbookServer server = new RunbookServer();
                    server.Name = _runbookServer;
                    job.RunbookServer = server;
                }


                // Add newly created job.
                _context.AddToJobs(job);

                if (_verbose)
                {
                    Console.WriteLine("\n\nStarting the job");
                }
                _context.SaveChanges();

                Console.WriteLine(string.Format("\n\nSuccessfully started job id: {0}", job.Id.ToString()));
                response = job.Id.ToString();
            }
            catch (DataServiceQueryException ex)
            {
                throw new ApplicationException("\n\nError starting the runbook.\n\n", ex);
            }


            return response;
        }

        public static Guid GetRunbookIdFromPath(OrchestratorWebService.OrchestratorContext context, string path)
        {
            Guid runbookId = new Guid();
            var retVal = context.Runbooks.Where(p => p.Path.Equals(path));
            if (retVal == null)
            {
                Console.WriteLine(string.Format("\n\nRunbook path '{0}' did not match any available runbooks.", path));
                return runbookId;
            }
            foreach (var runbook in retVal)
            {
                runbookId = new Guid(runbook.Id.ToString());
            }
            return runbookId;
        }

        public static string GetRunbookParameters()
        {
            Connect();

            if (_verbose)
            {
                Console.WriteLine("\n\nGetting runbook parameters from the web service.");
            }

            try
            {
                // Retrieve parameters for the runbook
                var runbookParams = _context.RunbookParameters.Where(runbookParam => runbookParam.RunbookId == _runbookID && runbookParam.Direction == "In");
                if (null == runbookParams)
                {
                    return "\n\nThe runbook does not require any parameters";
                }

                StringBuilder sb = new StringBuilder();
                foreach (var param in runbookParams)
                {
                    sb.AppendFormat("{0}='value';", param.Name.ToString());
                }

                if (_verbose)
                {
                    Console.WriteLine(string.Format("\n\nParameters for runbook ID '{0}'", _runbookID));
                }
                string parameterString = string.Format("\"{0}\"", sb.ToString().TrimEnd(';'));

                return parameterString;

            }
            catch (Exception e)
            {
                Console.WriteLine("\n\nError getting runbook parameters from the web service.\n\n" + e);
                return null;
            }
        }

        public static string FormatRunbookParameterString(OrchestratorWebService.OrchestratorContext context, Guid runbookID, Dictionary<string, string> inputParameters)
        {
            // Retrieve parameters for the runbook
            var runbookParams = context.RunbookParameters.Where(runbookParam => runbookParam.RunbookId == _runbookID && runbookParam.Direction == "In");
            if (runbookParams == null || runbookParams.Count() == 0)
            {
                return null;
            }
            if (null == inputParameters)
            {
                Console.WriteLine("\n\nStarting this runbook requires parameters but no parameters were specified.");
                throw new FormatException("Starting this runbook requires parameters but no parameters were specified.");
            }
            // Configure the XML for the parameters
            StringBuilder parametersXml = new StringBuilder();

            parametersXml.Append("<Data>");
            foreach (var param in runbookParams)
            {
                parametersXml.AppendFormat("<Parameter><ID>{0}</ID><Value>{1}</Value></Parameter>", param.Id.ToString("B"), _jobParameters[param.Name]);
            }
            parametersXml.Append("</Data>");

            if (_verbose)
            {
                Console.WriteLine("\n\nRunbook parameters XML:\n" + parametersXml.ToString());
            }
                
            return parametersXml.ToString();
        }

        static void PrintUsage()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("");
            sb.AppendLine(""); 
            sb.AppendLine("SCOJobRunner - a command line tool to invoke runbooks for Orchestrator 2012.");
            sb.AppendLine("");
            sb.AppendLine("SYNTAX");
            sb.AppendLine("ScoJobRunner -ID:<runbookID> [-Parameters:<job parameters>] ");
            sb.AppendLine(" [-Webserver:<server>] [-Port:<port number>] [-RunbookServer:<server>]");
            sb.AppendLine(" [[-Username:<username>] [-Domain:<domain>] [-Password:<password>]]");
            sb.AppendLine(" [-Verbose]");
            sb.AppendLine("");
            sb.AppendLine("ScoJobRunner -ID:<runbookID> -GetParameters [-Verbose]");
            sb.AppendLine("");
            sb.AppendLine("PARAMETERS");
            sb.AppendLine("");
            sb.AppendLine("  -ID <string>");
            sb.AppendLine("     The runbook ID to be invoked in the format of a GUID ");
            sb.AppendLine("     (i.e. \"2d2fd2b4-339a-47ad-9914-a217d9fd47a6\").");
            sb.AppendLine("");
            sb.AppendLine("  -Parameters <string>");
            sb.AppendLine("     A list of name-value pairs of the parameters needed by the runbook.");
            sb.AppendLine("     The format of these parameters is: \"Name=Value;Name=Value\"");
            sb.AppendLine("     Enclose in quotes to correctly handle spaces.");
            sb.AppendLine("     Example:  -Parameters \"Path='Path 1';Other Text='sample data here'\"");
            sb.AppendLine("");
            sb.AppendLine("  -GetParameters");
            sb.AppendLine("     When this parameter is added, the utility returns a list of the runbook");
            sb.AppendLine("     parameters formatted in a way you can use them on the command line.");
            sb.AppendLine("     Example return value: \"Path=Path 1;Other Text=sample data here\"");
            sb.AppendLine("");
            sb.AppendLine("");
            sb.AppendLine("  -Webserver <string>");
            sb.AppendLine("     The name of the Orchestrator server where the web service is installed.");
            sb.AppendLine("     If not provided, defaults to the current computer.");
            sb.AppendLine("");
            sb.AppendLine("  -Port <int>");
            sb.AppendLine("     The port number used by the web service. If not provided, defaults to 81.");
            sb.AppendLine("");
            sb.AppendLine("  -RunbookServer <string>");
            sb.AppendLine("     The Runbook Server where the job should be run.");
            sb.AppendLine("     If not provided, defaults to the configured primary Runbook Server.");
            sb.AppendLine("");
            sb.AppendLine("  -User <string> -Domain <string> -Password <string>");
            sb.AppendLine("     If provided, allows the use of alternate credentials for invoking the");
            sb.AppendLine("     runbook. All three parameters must be specified if used. ");
            sb.AppendLine("     If not provided, the current user credentials are used.");
            sb.AppendLine("");
            sb.AppendLine("  -Verbose");
            sb.AppendLine("     If this parameter is included, provides additional console output during");
            sb.AppendLine("     the processing of the job creation process.");
            sb.AppendLine("");
            sb.AppendLine("");
            Console.WriteLine(sb.ToString());
            
        }

    }


    /// <summary>
    /// Arguments parsing class
    /// </summary>
    public class Arguments
    {
        // Variables
        private StringDictionary Parameters;

        // Constructor
        public Arguments(string[] Args)
        {
            Parameters = new StringDictionary();
            Regex Splitter = new Regex(@"^-|:", RegexOptions.IgnoreCase | RegexOptions.Compiled);

            Regex Remover = new Regex(@"^['""]?(.*?)['""]?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

            string Parameter = null;
            string[] Parts;

            // Valid parameters forms:
            // -param{ ,:}((",')value(",'))
            // Examples: 
            // -param1 value1 -param2 -param3:"Foo;bar" 
            foreach (string Txt in Args)
            {
                Parts = Splitter.Split(Txt, 3);
                switch (Parts.Length)
                {
                    // Found a value (for the last parameter found (space separator))
                    case 1:
                        if (Parameter != null)
                        {
                            if (!Parameters.ContainsKey(Parameter))
                            {
                                Parts[0] =
                                    Remover.Replace(Parts[0], "$1");

                                Parameters.Add(Parameter, Parts[0]);
                            }
                            Parameter = null;
                        }
                        break;

                    // Found just a parameter
                    case 2:
                        // The last parameter is still waiting. 
                        // With no value, set it to true.
                        if (Parameter != null)
                        {
                            if (!Parameters.ContainsKey(Parameter))
                                Parameters.Add(Parameter, "true");
                        }
                        Parameter = Parts[1];
                        break;

                    // Parameter with enclosed value
                    case 3:
                        // The last parameter is still waiting. 
                        // With no value, set it to true.
                        if (Parameter != null)
                        {
                            if (!Parameters.ContainsKey(Parameter))
                                Parameters.Add(Parameter, "true");
                        }

                        Parameter = Parts[1];

                        // Remove possible enclosing characters (",')
                        if (!Parameters.ContainsKey(Parameter))
                        {
                            Parts[2] = Remover.Replace(Parts[2], "$1");
                            Parameters.Add(Parameter, Parts[2]);
                        }

                        Parameter = null;
                        break;
                }
            }
            // In case a parameter is still waiting
            if (Parameter != null)
            {
                if (!Parameters.ContainsKey(Parameter))
                    Parameters.Add(Parameter, "true");
            }
        }

        // Retrieve a parameter value if it exists (overriding C# indexer property)
        public string this[string Param]
        {
            get
            {
                return (Parameters[Param]);
            }
        }
    }

    
}
