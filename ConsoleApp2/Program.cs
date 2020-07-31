using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Activities.Statements;

namespace Dynamics365CRUD
{
    class Program
    {
        private const string trialAccount = "https://sscole5.crm.dynamics.com/";
        private const string trialAccountApi = "https://sscole5.api.crm.dynamics.com/";
        private const string password = "2nephi3120.D";
        private const string username = "ct@sscole5.onmicrosoft.com";
        private const string filename = "data.csv";
        static void Main(string[] args)
        {
            try
            {
                var connectionString = @"AuthType = Office365; Url = " + trialAccount + ";Username=" + username + ";Password=" + password;
                CrmServiceClient conn = new CrmServiceClient(connectionString);

                IOrganizationService service;
                service = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;

                // Get the file's text.
                string whole_file = System.IO.File.ReadAllText(filename);

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

                // TODO handle duplicates? worry about update a record?
                // For now, if exists already, just moves on, does no updating. If doesn't exist, creates, no trying to decide if mostly matches in other fields.

                // Read or Create contact
                // With Insurance Plan
                // Without Insurance Plan
                // As doctor
                // As minor
                Entity contact = new Entity("contact");

                // Read or Create dependencies
                Entity account = new Entity("account");
                Entity cole_insuranceplan2 = new Entity("cole_insuranceplan2");
                // account["address1_city"]
                // account["address1_line1"]
                // account["address1_line2"]
                // account["address1_postalcode"]
                // account["cole_carriertype"]
                // account["cole_insuranceparticipation"]
                // account["ct_accounttype"]
                // account["ct_healthcareprovidertype"]
                // account["description"]
                // account["name"]
                // account["numberofemployees"]
                // account["primarycontactid"]
                // account["revenue"]
                // account["telephone1"]
                
                contact["firstname"] = values[1,0];
                contact["lastname"] = values[1,1];
                contact["cole_birthday3"] = values[1,2];
                account["name"] = values[1,3];
                cole_insuranceplan2["cole_name"] = values[1,4];
                cole_insuranceplan2["cole_policynumber"] = values[1,5];
                cole_insuranceplan2["cole_groupnumber"] = values[1,6];
                contact["gendercode"] = values[1,7];
                contact["address1_line1"] = values[1,8];
                contact["emailaddress1"] = values[1,9];
                contact["telephone1"] = values[1,10];
                contact["cole_height"] = values[1,11];
                contact["cole_weight"] = values[1,12];
                contact["cole_contacttype"] = values[1,13];
                contact["cole_licensetype"] = values[1,14];
                contact["cole_specialization"] = values[1,15];
                account["name"] = values[1,16];

                Guid accountId = service.Create(account);
                Console.WriteLine("New account id: {0}.", accountId.ToString());
                Guid cole_insuranceplan2Id = service.Create(cole_insuranceplan2);
                Console.WriteLine("New cole_insuranceplan2 id: {0}.", cole_insuranceplan2Id.ToString());
                Guid contactId = service.Create(contact);
                Console.WriteLine("New contact id: {0}.", contactId.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
