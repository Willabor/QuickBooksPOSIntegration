using System;
using QBPOSFC4Lib; // Ensure this matches the actual namespace in your environment

namespace QuickBooksPOSIntegration
{
    public class QuickBooksPOSService
    {
        private QBPOSSessionManager _sessionManager;

        public QuickBooksPOSService()
        {
            _sessionManager = new QBPOSSessionManager();
        }

        public void OpenConnection(string appName)
        {
            _sessionManager.OpenConnection("", appName);
            // Adjust the session mode as needed (e.g., READ_WRITE if you plan to modify data)
            _sessionManager.BeginSession("", ENUM_QBPOS_SESSION_MODE.QBPOS_SM_READ_ONLY);
        }

        public void CloseConnection()
        {
            _sessionManager.EndSession();
            _sessionManager.CloseConnection();
        }

        public void FetchCustomers()
        {
            Console.WriteLine("Fetching customers...");

            // 1. Create the request message set
            // Note: The version numbers here (majorVersion=3, minorVersion=0) are just examples.
            //       Adjust them to match the QBPOS XML spec versions supported by your QBPOS installation.
            IMsgSetRequest requestMsgSet = _sessionManager.CreateMsgSetRequest("3.0");

            // Tell QB what to do if an error occurs: continue, stop, or return as-is.
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

            // 2. Append a Customer Query to the request
            ICustomerQuery customerQuery = requestMsgSet.AppendCustomerQueryRq();

            // OPTIONAL: set filters on the query (e.g., specific customer ListIDs, etc.)
            // customerQuery.ListIDList.Add("ABC-123");
            // customerQuery.LastName.SetValue("Smith");

            // 3. Send the request and get the response
            IMsgSetResponse responseMsgSet = _sessionManager.DoRequests(requestMsgSet);

            // 4. Process the response
            //    A single request can result in multiple responses. Here we only queried customers once,
            //    so there's likely only one response in the list.
            if (responseMsgSet != null && responseMsgSet.ResponseList != null)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response != null)
                {
                    // Always check status code, status message, etc.
                    // A status code of 0 typically indicates a successful request.
                    Console.WriteLine($"Status Code: {response.StatusCode}, Status Message: {response.StatusMessage}");

                    if (response.StatusCode == 0)
                    {
                        // 5. The response Detail should be an ICustomerRetList in this case.
                        ICustomerRetList customerRetList = response.Detail as ICustomerRetList;

                        if (customerRetList != null)
                        {
                            // 6. Iterate the customer list and access individual fields
                            for (int i = 0; i < customerRetList.Count; i++)
                            {
                                ICustomerRet customerRet = customerRetList.GetAt(i);

                                // Sample fields – adapt as needed for your use case:
                                string listID     = customerRet.ListID?.GetValue() ?? string.Empty;
                                string firstName  = customerRet.FirstName?.GetValue() ?? string.Empty;
                                string lastName   = customerRet.LastName?.GetValue() ?? string.Empty;
                                string fullName   = $"{firstName} {lastName}".Trim();
                                string phone      = customerRet.Phone?.GetValue() ?? string.Empty;
                                string email      = customerRet.Email?.GetValue() ?? string.Empty;

                                Console.WriteLine($"Customer: {fullName}, Phone: {phone}, Email: {email}, ListID: {listID}");
                                // You could store these in a list, database, etc.
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Customer query failed with status code: {response.StatusCode}");
                    }
                }
            }
        }
    }
}
