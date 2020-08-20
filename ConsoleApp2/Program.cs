using System;

using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using System.Globalization;
using Microsoft.Xrm.Sdk.Messages;
using System.IO;

namespace Dynamics365CRUD
{
    class Program
    {
        private const string credentialFile = @"c:\Users\colet\creds.txt";
        private const string dataFile = @"data.csv";
        static void Main(string[] args)
        {
            try
            {
                string trialAccount;
                string password;
                string username;
                using (StreamReader sr = File.OpenText(credentialFile))
                {
                    trialAccount = sr.ReadLine();
                    password = sr.ReadLine();
                    username = sr.ReadLine();
                }
                var connectionString = @"AuthType = Office365; Url = " + trialAccount + ";Username=" + username + ";Password=" + password;
                CrmServiceClient conn = new CrmServiceClient(connectionString);

                IOrganizationService service;
                service = (IOrganizationService)conn.OrganizationWebProxyClient ?? (IOrganizationService)conn.OrganizationServiceProxy;

                // Get the file's text.
                string whole_file = System.IO.File.ReadAllText(dataFile);

                // Split into lines.
                whole_file = whole_file.Replace('\n', '\r');
                string[] lines = whole_file.Split(new char[] { '\r' },
                    StringSplitOptions.RemoveEmptyEntries);

                // See how many rows and columns there are.
                int num_rows = lines.Length;
                int num_cols = lines[0].Split(',').Length;

                // Allocate the data array.
                string[,] values = new string[num_rows, num_cols];

                // Load the array.
                for (int r = 0; r < num_rows; r++)
                {
                    string[] line_r = lines[r].Split(',');
                    for (int c = 0; c < num_cols; c++)
                    {
                        values[r, c] = line_r[c];
                    }
                }
                
                // Use row to Create and Associate records
                for (var n = 1; n < num_rows; n++)
                {
                    Entity contact = new Entity("contact");
                    Entity account = new Entity("account");
                    Entity cole_insuranceplan2 = new Entity("cole_insuranceplan2");

                    Guid planId;
                    Guid contactId;
                    string accountName;

                    // Check if existing contact
                    var firstname = values[n, 0];
                    var lastname = values[n, 1];
                    var email = values[n, 9];
                    var fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='contact'><filter type='and'><condition attribute='emailaddress1' operator='eq' value='{email}' /><condition attribute='firstname' operator='eq' value='{firstname}' /><condition attribute='lastname' operator='eq' value='{lastname}' /></filter></entity></fetch>";
                    var res = conn.GetEntityDataByFetchSearchEC(fetchXML);
                    contactId = res.Entities.Count > 0 ? res.Entities[0].Id : Guid.Empty;

                    contact["cole_birthday3"] = DateTime.Parse(values[n, 2]);
                    contact["cole_weight"] = decimal.Parse(values[n, 12], CultureInfo.InvariantCulture);
                    contact["cole_height"] = int.Parse(values[n, 11]);
                    contact["gendercode"] = new OptionSetValue(int.Parse(values[n, 7]));
                    contact["firstname"] = values[n, 0];
                    contact["lastname"] = values[n, 1];
                    contact["address1_line1"] = values[n, 8];
                    contact["emailaddress1"] = values[n, 9];
                    contact["telephone1"] = values[n, 10];
                    
                    // Lookup to Insurance Plan (cole_insuranceplan) not required
                    
                    // Check for existing Plan
                    var policynumber = values[n, 5];
                    var fetchXMLPlan = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='";
                    fetchXMLPlan += @"cole_insuranceplan2'><filter type='and'><condition attribute='";
                    fetchXMLPlan += $@"cole_policynumber' operator='eq' value='{policynumber}' />";
                    fetchXMLPlan += @"</filter></entity></fetch>";
                    var resPlan = conn.GetEntityDataByFetchSearchEC(fetchXMLPlan);

                    // Create Insurance Plan (requires Contact ID, Acccount ID)
                    cole_insuranceplan2["cole_groupnumber"] = int.Parse(values[n, 6]);
                    cole_insuranceplan2["cole_name"] = values[n, 4];
                    cole_insuranceplan2["cole_policynumber"] = values[n, 5];
                    cole_insuranceplan2.Attributes.Add("cole_plantype", new OptionSetValue());
                    switch (values[n, 4])
                    {
                        case "G":
                            ((OptionSetValue)cole_insuranceplan2["cole_plantype"]).Value = 370510000;
                            break;
                        case "S":
                            ((OptionSetValue)cole_insuranceplan2["cole_plantype"]).Value = 370510001;
                            break;
                        case "B":
                            ((OptionSetValue)cole_insuranceplan2["cole_plantype"]).Value = 370510002;
                            break;
                        default:
                            throw new Exception();
                    }

                    // Lookup to Account.
                    // Lookup field is (cole_serviceproviderid) not required
                    accountName = values[n, 3];
                    var fetchXMLAccount = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='";
                    fetchXMLAccount += @"account'><filter type='and'><condition attribute='";
                    fetchXMLAccount += $@"name' operator='eq' value='{accountName}' />";
                    fetchXMLAccount += @"</filter></entity></fetch>";
                    var resAccount = conn.GetEntityDataByFetchSearchEC(fetchXMLAccount);
                    if (resAccount.Entities.Count == 0)
                    {
                        // Create Account
                        account["name"] = values[n, 3];
                        Guid accountId = service.Create(account);
                        Console.WriteLine("New account id: {0}. Name: {1}", accountId.ToString(), accountName);
                    }

                    if (resPlan.Entities.Count == 0)
                    {
                        planId = service.Create(cole_insuranceplan2);
                        Console.WriteLine("Plan built under Account: " + accountName);
                        Console.WriteLine("New cole_insuranceplan2 id: {0}.", planId.ToString());
                    }
                    else
                    {
                        planId = (Guid)resPlan.Entities[0].Id;
                        cole_insuranceplan2.Id = planId;
                        service.Update(cole_insuranceplan2);
                    }
                    contact.Attributes.Add("cole_insuranceplan", new EntityReference("cole_insuranceplan2", planId));

                    if (contactId == Guid.Empty)
                    {
                        // Create Contact
                        contactId = service.Create(contact);
                        Console.WriteLine("New contact id: {0}.", contactId.ToString());
                    }
                    else
                    {
                        contact.Id = contactId;

                        // Update will handle associations
                        service.Update(contact);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
