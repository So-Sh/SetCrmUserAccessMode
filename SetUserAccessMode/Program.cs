using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using Microsoft.Xrm.Sdk.Messages;

namespace SetUserAccessMode
{
    class Program
    {
        static void Main(string[] args)
        {
            var connection = CrmUtility.GetCRMConnection(CrmUtility.CRM_URL, CrmUtility.CRM_USERNAME, CrmUtility.CRM_PASSWORD);

            using (var _service = new OrganizationService(connection))
            {
                var businessUnits = ConfigurationManager.AppSettings["BU"].Split(',');

                Console.WriteLine("Connected to CRM: " + CrmUtility.CRM_URL);
                Console.WriteLine("--------------------------------------------------------");
                Console.WriteLine("Retrieveing user from following Business Unit(s): " + String.Join(",", businessUnits));

                QueryExpression queryExpression = new QueryExpression();
                queryExpression.EntityName = "systemuser";
                queryExpression.ColumnSet = new ColumnSet("accessmode", "businessunitid", "isdisabled");

                FilterExpression filter1 = queryExpression.Criteria.AddFilter(LogicalOperator.Or);
                for (int i = 0; i < businessUnits.Length; i++)
                {
                    filter1.Conditions.Add(new ConditionExpression("businessunitid", ConditionOperator.Equal,
                        CrmUtility.GetUserBusinessUnit(connection, businessUnits[i]).Id));                    
                }


                ConditionExpression conditionExpression2 = new ConditionExpression();
                conditionExpression2.AttributeName = "isdisabled";
                conditionExpression2.Values.Add(false);
                conditionExpression2.Operator = ConditionOperator.Equal;
                
               
                queryExpression.Criteria.AddCondition(conditionExpression2);
                queryExpression.Criteria.FilterOperator = LogicalOperator.And;

                EntityCollection entityColl = _service.RetrieveMultiple(queryExpression);

                // access mode ==> 0 --> Read Write // 1 --> Adminstrative // Read Only --> 2

                var R_W_Users = entityColl.Entities.Where(e => ((OptionSetValue)e.Attributes["accessmode"]).Value == 0).
                    Select(e => e.ToEntity<Entity>()).ToList<Entity>();

                var ROnly_Users = entityColl.Entities.Where(e => ((OptionSetValue)e.Attributes["accessmode"]).Value == 2).
                                Select(e => e.ToEntity<Entity>()).ToList<Entity>();


                Console.WriteLine("");
                Console.WriteLine("{0} users with Read/Write Access", R_W_Users.Count);
                Console.WriteLine("{0} users with Read-Only Access", ROnly_Users.Count);
                Console.WriteLine("--------------------------------------------------------");

                Console.WriteLine("Type 1 to set users to Read-Only mode, type 2 to set users to Read/Write mode and hit enter");

                //Let the user select what to do
                string opCode = Console.ReadLine();
                while(opCode != "1" && opCode != "2")
                {
                    Console.WriteLine("Please select a valid option and hit enter: 1 to set users to Read-Only mode, 2 to set users to Read/Write mode");
                    opCode = Console.ReadLine();
                }               

                if (opCode == "1")
                {
                    Console.WriteLine("Do you want to set all {0} to Read-Only mode?", R_W_Users.Count);
                    Console.WriteLine("Select an option and hit enter: y= Yes, n =  No");
                    string answer = Console.ReadLine();
                    if (answer.ToLower() == "y")
                    {
                        UpdateUserAccessMode(R_W_Users,_service,2);
                    }
                }
                else if (opCode == "2")
                {
                    Console.WriteLine("Do you want to set all {0} to Read/Write mode?", ROnly_Users.Count);
                    Console.WriteLine("Select an option and hit enter: y= Yes, n =  No");
                    string answer = Console.ReadLine();
                    if (answer.ToLower() == "y")
                    {
                        UpdateUserAccessMode(ROnly_Users, _service, 0);
                    }
                }

                Console.WriteLine("");
                Console.WriteLine("End of program");

                Console.ReadLine();
                return;
            }
        }

        private static void UpdateUserAccessMode(List<Entity> users, OrganizationService _service, int accessmode)
        {
            // Create an ExecuteMultipleRequest object.
            var requestWithResults = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses. 
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()
            };

            foreach (var user in users)
            {
                UpdateRequest updateRequest = new UpdateRequest() { Target = user };
                user.Attributes["accessmode"] = new OptionSetValue(accessmode);
                requestWithResults.Requests.Add(updateRequest);
            }
            Console.WriteLine("");
            Console.WriteLine("Updating Access Mode for {0} users...", requestWithResults.Requests.Count);
            // Execute all the requests in the request collection using a single web method call.
            ExecuteMultipleResponse responseWithResults =
                (ExecuteMultipleResponse)_service.Execute(requestWithResults);

            // Display the results returned in the responses.
            var success = responseWithResults.Responses.Where(r => r.Fault == null).
                Select(r => r.Response).ToList<OrganizationResponse>();

            var fail = responseWithResults.Responses.Where(r => r.Fault != null).
                Select(r => r.Response).ToList<OrganizationResponse>();

            Console.WriteLine("");
            Console.WriteLine("Successfully updated {0} users", success.Count);
            Console.WriteLine("");
            Console.WriteLine("Failed to update {0} users", fail.Count);
        }
    }
}
