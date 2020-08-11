using System;

using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using System.Globalization;
using Microsoft.Xrm.Sdk.Messages;

namespace Dynamics365CRUD
{
    class Program
    {
        private const string trialAccount = "https://sscole5.crm.dynamics.com/";
        private const string trialAccountApi = "https://sscole5.api.crm.dynamics.com/";
        private const string password = "2nephi3120.D";
        private const string username = "ct@sscole5.onmicrosoft.com";
        private const string filename = @"data.csv";
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

                for (var n = 1; n < num_rows; n++)
                {

                    // Read or Create contact
                    // With Insurance Plan
                    // Without Insurance Plan
                    // As doctor
                    // As minor
                    Entity contact = new Entity("contact");
                    Entity account = new Entity("account");
                    Entity cole_insuranceplan2 = new Entity("cole_insuranceplan2");

                    Guid planId;
                    //Guid accountId;
                    string accountName;
                    var firstname = values[n, 0];
                    var lastname = values[n, 1];
                    var email = values[n, 9];
                    var fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='contact'><filter type='and'><condition attribute='emailaddress1' operator='eq' value='{email}' /><condition attribute='firstname' operator='eq' value='{firstname}' /><condition attribute='lastname' operator='eq' value='{lastname}' /></filter></entity></fetch>";
                    var res = conn.GetEntityDataByFetchSearchEC(fetchXML);
                    if (res.Entities.Count == 0)
                    {
                        // Create Contact (Requires Insurance Plan ID)
                        var policynumber = values[n, 5];
                        var fetchXMLPlan = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='";
                        fetchXMLPlan += @"cole_insuranceplan2'><filter type='and'><condition attribute='";
                        fetchXMLPlan += $@"cole_policynumber' operator='eq' value='{policynumber}' />";
                        fetchXMLPlan += @"</filter></entity></fetch>";
                        var resPlan = conn.GetEntityDataByFetchSearchEC(fetchXMLPlan);
                        if (resPlan.Entities.Count == 0)
                        {
                            // Create Insurance Plan (Requires Account Name)
                            // TODO find better way to check existing Account.
                            // FOR NOW see if account exists. All I have from data.csv is the account name.
                            accountName = values[n, 3];
                            var fetchXMLAccount = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='";
                            fetchXMLAccount += @"account'><filter type='and'><condition attribute='";
                            fetchXMLAccount += $@"name' operator='eq' value='{accountName}' />";
                            fetchXMLAccount += @"</filter></entity></fetch>";
                            var resAccount = conn.GetEntityDataByFetchSearchEC(fetchXMLAccount);
                            if (resAccount.Entities.Count == 0)
                            {
                                // Create Account
                                Console.WriteLine(values[n, 3]); account["name"] = values[n, 3];
                                Guid accountId = service.Create(account);
                                Console.WriteLine("New account id: {0}. Name: {1}", accountId.ToString(), accountName);
                                // TODO add all data to account.
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
                            }
                            //Console.WriteLine(values[n, 6]);
                            cole_insuranceplan2["cole_groupnumber"] = int.Parse(values[n, 6]);
                            //Console.WriteLine(values[n, 4]);
                            cole_insuranceplan2["cole_name"] = values[n, 4];
                            //Console.WriteLine(values[n, 5]);
                            cole_insuranceplan2["cole_policynumber"] = values[n, 5];
                            //Console.WriteLine("Type of cole_insuranceplan2: {0}", cole_insuranceplan2.GetType());
                            //Console.WriteLine("Type of cole_insuranceplan2 Attributes: {0}", cole_insuranceplan2.Attributes); //Microsoft.Xrm.Sdk.AttributeCollection
                            //Console.WriteLine(cole_insuranceplan2.Attributes.Keys); //System.Collections.Generic.Dictionary`2+KeyCollection[System.String,System.Object]
                            //var lookatme = cole_insuranceplan2.Attributes.Keys; // just has the three keys I already used, not all the entity's keys.
                            //foreach(var a in cole_insuranceplan2.Attributes)
                            //{
                            //    Console.WriteLine(a.ToString()); // [cole_groupnumber, 1] etc.
                            //}
                            cole_insuranceplan2.Attributes.Add("cole_plantype", new OptionSetValue());
                            switch (values[n, 4])
                            {
                                case "G":
                                    //Console.WriteLine("G type is: ");
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
                            planId = service.Create(cole_insuranceplan2);
                            Console.WriteLine("Plan built under Account: " + accountName);
                            Console.WriteLine("New cole_insuranceplan2 id: {0}.", planId.ToString());
                        }
                        else
                        {
                            planId = (Guid)resPlan.Entities[0]["Id"];
                        }
                        // Create
                        Console.WriteLine(values[n, 2]); contact["cole_birthday3"] = DateTime.Parse(values[n, 2]);
                        Console.WriteLine(values[n, 12]); contact["cole_weight"] = decimal.Parse(values[n, 12], CultureInfo.InvariantCulture);
                        Console.WriteLine(values[n, 11]); contact["cole_height"] = int.Parse(values[n, 11]);
                        Console.WriteLine(values[n, 7]); contact["gendercode"] = new OptionSetValue(int.Parse(values[n, 7]));
                        //Console.WriteLine(values[n, 13]); contact["cole_contacttype"] = new OptionSetValue(int.Parse(values[n, 13]));
                        //Console.WriteLine(values[n, 14]); contact["cole_licensetype"] = new OptionSetValue(int.Parse(values[n, 14]));
                        Console.WriteLine(values[n, 0]); contact["firstname"] = values[n, 0];
                        Console.WriteLine(values[n, 1]); contact["lastname"] = values[n, 1];
                        Console.WriteLine(values[n, 8]); contact["address1_line1"] = values[n, 8];
                        Console.WriteLine(values[n, 9]); contact["emailaddress1"] = values[n, 9];
                        Console.WriteLine(values[n, 10]); contact["telephone1"] = values[n, 10];
                        //Console.WriteLine(values[n, 15]); contact["cole_specialization"] = values[n, 15];

                        // make sure lookup fields are correct, use GUIds from Plan and Account
                        //contact["cole_serviceproviderid"] = new EntityReference("account", accountId)
                        // ref t oplan ? contact["cole_serviceproviderid"] = new EntityReference("account", planId)

                        Guid contactId = service.Create(contact);
                        Console.WriteLine("New contact id: {0}.", contactId.ToString());

                        // TODO update insurance plan with patient
                    }
                    // if you can find the contact try to update it.
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
