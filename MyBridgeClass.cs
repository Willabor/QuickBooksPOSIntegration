using System;
using System.Threading.Tasks;
using Interop.QBPOSFC4.Full; // Ensure this matches the namespace in your Interop DLL

namespace QuickBooksPOSIntegration
{
    public class MyBridgeClass
    {
        // Single session manager for the entire session
        private QBPOSSessionManagerClass _sessionManager;

        // Define your custom enum for QuickBooks POS session modes
        public enum MySessionMode
        {
            qbposSMExclusive = 0,  // For exclusive mode
            qbposSMReadOnly  = 1   // For read-only mode
        }

        // Edge.js typically calls "Invoke" and passes in a dynamic object (input).
        public async Task<object> Invoke(dynamic input)
        {
            // Placeholder for real async work
            await Task.CompletedTask;

            string operation = (string)input.operation;

            switch (operation)
            {
                case "OpenConnection":
                    return OpenConnection((string)input.appName);

                case "FetchCustomers":
                    return FetchCustomers();

                case "CloseConnection":
                    return CloseConnection();

                default:
                    return $"Unknown operation: {operation}";
            }
        }

        /// <summary>
        /// Opens the connection to QuickBooks POS and begins a read-only session.
        /// </summary>
        private string OpenConnection(string appName)
        {
            try
            {
                // Create a new QBPOSSessionManagerClass instance
                _sessionManager = new QBPOSSessionManagerClass();

                // Open the connection to QuickBooks POS
                // (Unlike QuickBooks Desktop, POS doesn't require a data file path.)
                _sessionManager.OpenConnection("", appName);

                // Begin the POS session in READ-ONLY mode
                // If you need exclusive mode, change to: MySessionMode.qbposSMExclusive
                _sessionManager.BeginSession(MySessionMode.qbposSMReadOnly.ToString());

                return "Connection opened successfully (read-only).";
            }
            catch (Exception ex)
            {
                return $"Error opening connection: {ex.Message}";
            }
        }

        /// <summary>
        /// Example method to fetch customers from QuickBooks POS.
        /// </summary>
        private string FetchCustomers()
        {
            try
            {
                if (_sessionManager == null)
                {
                    return "Session manager is not initialized. Please open a connection first.";
                }

                // Create a simple Customer Query request
                // Adjust the QBXML version if needed (3, 0) might be too old; check POS docs
                IMsgSetRequest requestMsgSet = _sessionManager.CreateMsgSetRequest(3, 0);
                requestMsgSet.AppendCustomerQueryRq();

                // Send the request and get the response as QBXML
                string qbxmlResponse = _sessionManager.DoRequests(requestMsgSet).ToXMLString();
                return qbxmlResponse; // Return the raw QBXML response
            }
            catch (Exception ex)
            {
                return $"Error fetching customers: {ex.Message}";
            }
        }

        /// <summary>
        /// Closes the session and connection to QuickBooks POS.
        /// </summary>
        private string CloseConnection()
        {
            try
            {
                if (_sessionManager != null)
                {
                    _sessionManager.EndSession();
                    _sessionManager.CloseConnection();
                    _sessionManager = null;
                }
                return "Connection closed successfully.";
            }
            catch (Exception ex)
            {
                return $"Error closing connection: {ex.Message}";
            }
        }
    }
}
