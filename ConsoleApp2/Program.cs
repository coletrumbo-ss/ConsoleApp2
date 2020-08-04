using System;

using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using System.Globalization;

namespace Dynamics365CRUD
{
    class Program
    {
        private const string trialAccount = "https://sscole5.crm.dynamics.com/";
        private const string trialAccountApi = "https://sscole5.api.crm.dynamics.com/";
        private const string password = "2nephi3120.D";
        private const string username = "ct@sscole5.onmicrosoft.com";
        private const string filename = @"C:\Users\colet\Downloads\data.csv";
        static void Main(string[] args)
        {
            try
            {
                var connectionString = @"AuthType = Office365; Url = " + trialAccount + ";Username=" + username + ";Password=" + password;
                CrmServiceClient conn = new CrmServiceClient(connectionString);

                IOrganizationService service;
                service = (IOrganizationService)conn.OrganizationWebProxyClient ?? (IOrganizationService)conn.OrganizationServiceProxy;

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

                for(var n = 1; n < num_rows; n++)
                {
                    Console.WriteLine(values[n, 2]); contact["cole_birthday3"] = DateTime.Parse(values[n, 2]);
                    Console.WriteLine(values[n, 12]); contact["cole_weight"] = decimal.Parse(values[n, 12], CultureInfo.InvariantCulture);
                    Console.WriteLine(values[n, 6]); cole_insuranceplan2["cole_groupnumber"] = int.Parse(values[n, 6]);
                    Console.WriteLine(values[n, 11]); contact["cole_height"] = int.Parse(values[n, 11]);
                    Console.WriteLine(values[n, 7]); contact["gendercode"] = new OptionSetValue(int.Parse(values[n, 7]));
                    //Console.WriteLine(values[n, 13]); contact["cole_contacttype"] = new OptionSetValue(int.Parse(values[n, 13]));
                    //if(values[n, 13] == "patient")
                    //{
                    //    Console.WriteLine(values[n, 14]); contact["cole_licensetype"] = new OptionSetValue(int.Parse(values[n, 14]));
                    //}
                    Console.WriteLine(values[n, 0]); contact["firstname"] = values[n, 0];
                    Console.WriteLine(values[n, 1]); contact["lastname"] = values[n, 1];
                    Console.WriteLine(values[n, 3]); account["name"] = values[n, 3];
                    Console.WriteLine(values[n, 4]); cole_insuranceplan2["cole_name"] = values[n, 4];
                    Console.WriteLine(values[n, 5]); cole_insuranceplan2["cole_policynumber"] = values[n, 5];
                    Console.WriteLine(values[n, 8]); contact["address1_line1"] = values[n, 8];
                    Console.WriteLine(values[n, 9]); contact["emailaddress1"] = values[n, 9];
                    Console.WriteLine(values[n, 10]); contact["telephone1"] = values[n, 10];
                    Console.WriteLine(values[n, 15]); contact["cole_specialization"] = values[n, 15];
                    Console.WriteLine(values[n, 16]); account["name"] = values[n, 16];

                    Guid accountId = service.Create(account);
                    Console.WriteLine("New account id: {0}.", accountId.ToString());
                    Guid cole_insuranceplan2Id = service.Create(cole_insuranceplan2);
                    Console.WriteLine("New cole_insuranceplan2 id: {0}.", cole_insuranceplan2Id.ToString());
                    Guid contactId = service.Create(contact);
                    Console.WriteLine("New contact id: {0}.", contactId.ToString());
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
