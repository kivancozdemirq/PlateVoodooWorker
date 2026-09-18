namespace PlateVoodooWorker
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Worker worker = new Worker();

            int operationId = 2; // OperationId
            int caseState = 1;   // Excel okuma ve DB yazma (Case 1) - OSR analiz ve API çağrısı (Case 2) 

            Console.WriteLine("Worker başlatılıyor...");
            worker.ExecuteProcess(operationId, caseState);
            Console.WriteLine("Case tamamlandı.");

            // IHost host = Host.CreateDefaultBuilder(args)
            //     .ConfigureServices(services =>
            //     {
            //         services.AddHostedService<Worker>();
            //     })
            //     .Build();

            // host.Run();
        }
    }
}