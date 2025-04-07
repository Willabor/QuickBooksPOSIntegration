using System;
using Interop.QBPOSFC4.Full;

namespace QuickBooksPOSIntegration
{
    public enum ENUM_QBPOS_SESSION_MODE
    {
        QBPOS_SM_READ_ONLY = 0,
        QBPOS_SM_READ_WRITE = 1
    }

    public class QuickBooksPOSService
    {
        private QBPOSSessionManager _sessionManager;

        public QuickBooksPOSService()
        {
            // Instantiate the COM class that implements the QBPOSSessionManager interface
            _sessionManager = new QBPOSSessionManagerClass();
        }

        public void OpenConnection(string appName)
        {
            try
            {
                // Open the connection to QuickBooks POS
                _sessionManager.OpenConnection("", appName);

                // Begin the session (default mode is read-only)
                _sessionManager.BeginSession(""); // Removed the second argument
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening connection: {ex.Message}");
                throw;
            }
        }

        public void CloseConnection()
        {
            try
            {
                // End the session and close the connection
                _sessionManager.EndSession();
                _sessionManager.CloseConnection();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error closing connection: {ex.Message}");
                throw;
            }
        }

        public void FetchCustomers()
        {
            Console.WriteLine("Fetching customers...");

            // 1. Create the request message set
            // Note: The version numbers here (majorVersion=3, minorVersion=0) are just examples.
            //       Adjust them to match the QBPOS XML spec versions supported by your QBPOS installation.
            IMsgSetRequest requestMsgSet = _sessionManager.CreateMsgSetRequest(3, 0);

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
                                // Email property is not available in ICustomerRet; removing email-related code
                                Console.WriteLine($"Customer: {fullName}, Phone: {phone}, ListID: {listID}");
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
