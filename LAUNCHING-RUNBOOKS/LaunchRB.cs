private static Runbook GetRunbookByPath(string serviceURL, string path)
{
    Runbook runbook = null;

    SCOService.OrchestratorContext context =
       new SCOService.OrchestratorContext(new Uri(serviceURL));

    // Set credentials to default or a specific user.
    context.Credentials = System.Net.CredentialCache.DefaultCredentials;

    // Get the runbook based on the entered path
    runbook = context.Runbooks.Where(rb => rb.Path == path).First();
    return runbook;
}

private static List<RunbookParameter> GetRunbookParametersById(string serviceURL,
   Guid runbookID)
{
    List<RunbookParameter> parameters = null;
    SCOService.OrchestratorContext context =
       new SCOService.OrchestratorContext(new Uri(serviceURL));

    // Set credentials to default or a specific user.
    context.Credentials = System.Net.CredentialCache.DefaultCredentials;

    // Get the parameters based on the runbook ID
    parameters = context.RunbookParameters.Where(rp => rp.RunbookId == runbookID
       && rp.Direction == "In").ToList();
    return parameters;
}

private static Job StartRunbookJob(string serviceURL, Guid runbookID,
   List<RunbookParameter> runbookParameters, Hashtable runbookParameterValues)
{
    Job job = new Job();
    // Configure the XML for the parameters
    StringBuilder parametersXml = new StringBuilder();
    if (runbookParameters != null && runbookParameters.Count() > 0)
    {
        parametersXml.Append("<Data>");
        foreach (var param in runbookParameters)
        {
           parametersXml.AppendFormat(
              "<Parameter><ID>{0}</ID><Value>{1}</Value></Parameter>",
              param.Id.ToString("B"), runbookParameterValues[param.Name]);
        }
        parametersXml.Append("</Data>");
    }

    try
    {
        // Create new job and assign runbook Id and parameters.
        job.RunbookId = runbookID;
        job.Parameters = parametersXml.ToString();

        SCOService.OrchestratorContext context =
           new SCOService.OrchestratorContext(new Uri(serviceURL));

        // Set credentials to default or a specific user.
        context.Credentials = System.Net.CredentialCache.DefaultCredentials;

        // Add newly created job.
        context.AddToJobs(job);
        context.SaveChanges();

        return job;
    }
    catch (DataServiceQueryException ex)
    {
        throw new ApplicationException("Error starting runbook.", ex);
    }
}

// main program

static void Main(string[] args)
{
    // Begin change values
    string serviceURL = "http://orch2012sp1:81/Orchestrator2012/Orchestrator.svc";
    string runbookPath = @"\DCS - Automation Library\Component Runbooks\SystemCenter\Virtual Machine Manager\1.0\Get VMM User Role";
    Hashtable parameters = new Hashtable();
    parameters["VMM Server Name"] = "VMM2012SP1.dcs.corp";
    parameters["User Role Name"] = "TenantAdmin1";
    // End change values

    Runbook runbook = GetRunbookByPath(serviceURL, runbookPath);
    List<RunbookParameter> runbookParameters = GetRunbookParametersById(serviceURL,
                                                   runbook.Id);
    Job job = StartRunbookJob(serviceURL, runbook.Id, runbookParameters, parameters);
    Console.WriteLine("Successfully started runbook. Job ID: {0}", job.Id.ToString());
    Console.ReadKey();
}