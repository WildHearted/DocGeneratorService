using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;

namespace DocGeneratorService
	{
	static class Program
		{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		static void Main()
			{

			ServiceBase[] objServicesBaseToRun;
			objServicesBaseToRun = new ServiceBase[]
				{
				new DocGeneratorServiceBase()
				};
		
				ServiceBase.Run(objServicesBaseToRun);
		
			}
		//static void RunInteractive(ServiceBase[] parServiceBaseToRun)
		//	{
		//	Console.WriteLine("Services running in INTERACTIVE mode.\n");

		//	MethodInfo onStartMethod = typeof(ServiceBase).GetMethod(
		//		name: "OnStart",
		//		bindingAttr: System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);

		//	foreach(ServiceBase objServiceBase in parServiceBaseToRun)
		//		{
		//		Console.Write("\tStarting {0}...", objServiceBase.ServiceName);
		//		onStartMethod.Invoke(objServiceBase, new object[] { new string[] { } });
		//		Console.WriteLine("Started");
		//		}

		//	Console.WriteLine("\n\n");

		//	Console.WriteLine("Press any key to stop the service and end the process...");
		//	Console.ReadKey();
		//	Console.WriteLine();

		//	MethodInfo onStopMethod = typeof(ServiceBase).GetMethod(
		//		name: "OnStop",
		//		bindingAttr: BindingFlags.Instance | BindingFlags.NonPublic);

		//	foreach(ServiceBase objServiceBase in parServiceBaseToRun)
		//		{
		//		Console.Write("Stopping {0}...", objServiceBase.ServiceName);
		//		onStopMethod.Invoke(objServiceBase, null);
		//		Console.Write("Stopped");
		//		}

		//	Console.WriteLine("All Services Stopped.");

		//	Thread.Sleep(1000);
		//	}
		}
	}

