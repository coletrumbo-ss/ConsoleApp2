using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Carl.Dynamics365CRUD
{
    class Program
    {
        private const string trialAccount = "https://sscole5.crm.dynamics.com/";
        private const string trialAccountApi = "https://sscole5.api.crm.dynamics.com/";
        private const string password = "2nephi3120.D";
        private const string username = "ct@sscole5.onmicrosoft.com";
        static void Main(string[] args)
        {
            try
            {

                var connectionString = @"AuthType = Office365; Url = " + trialAccount + ";Username=" + username + ";Password=" + password;
                CrmServiceClient conn = new CrmServiceClient(connectionString);

                IOrganizationService service;
                service = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;

                // Create a new record
                Entity contact = new Entity("contact");
                contact["firstname"] = "Bob";
                contact["lastname"] = "Smith";
                Guid contactId = service.Create(contact);
                Console.WriteLine("New contact id: {0}.", contactId.ToString());

                // Retrieve a record using Id
                Entity retrievedContact = service.Retrieve(contact.LogicalName, contactId, new ColumnSet(true));
                Console.WriteLine("Record retrieved {0}", retrievedContact.Id.ToString());

                // Update record using Id, retrieve all attributes
                Entity updatedContact = new Entity("contact");
                updatedContact = service.Retrieve(contact.LogicalName, contactId, new ColumnSet(true));
                updatedContact["jobtitle"] = "CEO";
                updatedContact["emailaddress1"] = "test@test.com";
                service.Update(updatedContact);
                Console.WriteLine("Updated contact");

                // Retrieve specific fields using ColumnSet
                ColumnSet attributes = new ColumnSet(new string[] { "jobtitle", "emailaddress1" });
                retrievedContact = service.Retrieve(contact.LogicalName, contactId, attributes);
                foreach (var a in retrievedContact.Attributes)
                {
                    Console.WriteLine("Retrieved contact field {0} - {1}", a.Key, a.Value);
                }

                // Delete a record using Id
                service.Delete(contact.LogicalName, contactId);
                Console.WriteLine("Deleted");
                Console.ReadLine();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
