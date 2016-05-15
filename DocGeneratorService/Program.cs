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
			ServiceBase[] objServicesBase;
			objServicesBase = new ServiceBase[]
				{
				new DocGeneratorServiceBase()
				};
		
			ServiceBase.Run(objServicesBase);

			}
		}
	}

