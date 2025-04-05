using System;

namespace QuickBooksPOSIntegration
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var service = new QuickBooksPOSService();
            service.OpenConnection("MyTestApp");
            service.FetchCustomers();
            service.CloseConnection();
        }
    }
}