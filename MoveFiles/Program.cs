using System.ServiceProcess;
using System.Threading;

namespace MoveFiles
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new MoveFile()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}