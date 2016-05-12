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
			Console.WriteLine("Main module of DocGeneratorService - Begin @ " + DateTime.Now.ToString("G"));
			AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

			ServiceBase[] objServicesBase;
			objServicesBase = new ServiceBase[]
				{
				new DocGeneratorServiceBase()
				};
		
			ServiceBase.Run(objServicesBase);

			Console.WriteLine("Main module of DocGeneratorService - Ended @ " + DateTime.Now.ToString("G"));
			}

		private static void CurrentDomain_UnhandledException(Object sender, UnhandledExceptionEventArgs e)
			{
			if(e != null && e.ExceptionObject != null)
				{
				Console.WriteLine(DateTime.Now.ToString("G") + " +++ Unexpected Exception occurred +++");
				//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " +++ Exception occurred in OnStart service +++");
				}
			}
		}
	}

