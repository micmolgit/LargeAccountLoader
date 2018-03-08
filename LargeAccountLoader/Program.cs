using System;
using System.ServiceModel;
using Nito.AsyncEx;

namespace LargeAccountLoader
{
    class Program
    {
        static public void Main(string[] args)
        {
            AsyncContext.Run(() => MainAsync(args));
        }

        static async void MainAsync(string[] args)
        {
            var largeAccountLoader = new LargeAccountLoader();

            try
            {
                var accountGuid = "78E89CBF-AEB9-E611-80FC-005056A2D11D";
                var cptEntities = await largeAccountLoader.Load(accountGuid);

                Console.WriteLine($"{cptEntities} entities were found for {accountGuid}");
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Console.WriteLine("The application terminated with an error.");
                Console.WriteLine($"Timestamp: { ex.Detail.Timestamp}");
                Console.WriteLine($"Code: {ex.Detail.ErrorCode}");
                Console.WriteLine($"Message: {ex.Detail.Message}");
                Console.WriteLine($"Plugin Trace: {ex.Detail.TraceText}");
                Console.WriteLine("Inner Fault: {0}",
                    null == ex.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
            }
            catch (System.TimeoutException ex)
            {
                Console.WriteLine("The application terminated with an error.");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                Console.WriteLine("Inner Fault: {0}",
                ex.InnerException.Message ?? "No Inner Fault");
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("The application terminated with an error.");
                Console.WriteLine(ex.Message);

                // Display the details of the inner exception.
                if (ex.InnerException != null)
                {
                    Console.WriteLine(ex.InnerException.Message);

                    if (ex.InnerException is FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> fe)
                    {
                        Console.WriteLine($"Timestamp: {fe.Detail.Timestamp}");
                        Console.WriteLine($"Code: {fe.Detail.ErrorCode}");
                        Console.WriteLine($"Message: {fe.Detail.Message}");
                        Console.WriteLine($"Plugin Trace: {fe.Detail.TraceText}");
                        Console.WriteLine("Inner Fault: {0}",
                            null == fe.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
                    }
                }
            }
            finally
            {
                largeAccountLoader.Terminate();
            }
        }
    }
}
