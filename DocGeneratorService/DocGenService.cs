using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocGeneratorCore;

namespace DocGeneratorService
	{
	public partial class DocGenService:ServiceBase
		{
		public enum ServiceState
			{
			SERVICE_STOPPED = 0x00000001,
			SERVICE_START_PENDING = 0x00000002,
			SERVICE_STOP_PENDING = 0x00000003,
			SERVICE_RUNNING = 0x00000004,
			SERVICE_CONTINUE_PENDING = 0x00000005,
			SERVICE_PAUSE_PENDING = 0x00000006,
			SERVICE_PAUSED = 0x00000007,
			}
		[StructLayout(LayoutKind.Sequential)]
		public struct ServiceStatus
			{
			public long dwServiceType;
			public ServiceState dwCurrentState;
			public long dwControlsAccepted;
			public long dwWin32ExitCode;
			public long dwServiceSpecificExitCode;
			public long dwCheckPoint;
			public long dwWaitHint;
			};

		[DllImport("advapi32.dll", SetLastError = true)]
		private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);
	
		public DocGenService()
			{
			InitializeComponent();
			objEventLog = new System.Diagnostics.EventLog();
			if(!System.Diagnostics.EventLog.SourceExists("DocGeneratorEventSource"))
				{
				System.Diagnostics.EventLog.CreateEventSource(source: "DocGeneratorEventSource", logName: "DocGeneratorLog");
				}
			objEventLog.Source = "DocGeneratorEventSource";
			objEventLog.Log = "DocGeneratorLog";
			}

		protected override void OnStart(string[] args)
			{
			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " +++ Starting DocGenerator Service +++");

			// Update the service state to Start Pending. 
			ServiceStatus objServiceStatus = new ServiceStatus();
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
			objServiceStatus.dwWaitHint = 100000; // 10 seconds
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			System.Timers.Timer objTimer = new System.Timers.Timer();
			objTimer.Interval = 60000; // timer fire every 60 seconds.
			// Inistiate the Timer's Event Ticker
			objTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.objTimer_Tick);
			// Switch the envent on
			objTimer.Enabled = true;

			// Update the service state to Running. 
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " +++ Started DocGenerator Service +++");
			}

		protected override void OnStop()
			{
			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " --- Starting DocGenerator Service ---");
			
			// Update the service state to Stop Pending. 
			ServiceStatus objServiceStatus = new ServiceStatus();
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
			objServiceStatus.dwWaitHint = 3000000; // 300 seconds - 5 minutes
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objTimer.Enabled = false;

			// Update the service state to Stopped. 
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " --- DocGenerator Service Started ---");
			}

		protected override void OnPause()
			{
			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === Pausing DocGenerator Service ===");

			// Update the service state to Pause Pending. 
			ServiceStatus objServiceStatus = new ServiceStatus();
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_PAUSE_PENDING;
			objServiceStatus.dwWaitHint = 3000000; // 300 seconds - 5 minutes
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objTimer.Enabled = false;

			// Update the service state to Stopped. 
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_PAUSED;
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === DocGenerator Service Paused ===");
			}

		protected override void OnContinue()
			{
			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === Continuing DocGenerator Service ===");

			// Update the service state to Start Pending. 
			ServiceStatus objServiceStatus = new ServiceStatus();
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_CONTINUE_PENDING;
			objServiceStatus.dwWaitHint = 100000; // 10 seconds
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objTimer.Enabled = true;

			// Update the service state to Running. 
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === DocGenerator Service Continued ===");
			}

		protected override void OnShutdown()
			{
			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === Shutting Down DocGenerator Service ===");

			// Update the service state to Shatdown Pending. 
			ServiceStatus objServiceStatus = new ServiceStatus();
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
			objServiceStatus.dwWaitHint = 3000000; // 300 seconds - 5 minutes
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objTimer.Enabled = false;

			// Update the service state to Stopped. 
			objServiceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
			SetServiceStatus(this.ServiceHandle, ref objServiceStatus);
			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === DocGenerator Service Stopped/Shutdown ===");
			}


		private void objTimer_Tick(object sender, EventArgs e)
			{
			objEventLog.WriteEntry(DateTime.Now.ToString("G") + " +++ Started DocGenerator Service +++");

			// Invoke the DocGeneratorCore's MainController object MainProcess method
			DocGeneratorCore.MainController objMainController = new MainController();
			objMainController.MainProcess();

			objEventLog.WriteEntry("     + Event Ended... @ " + DateTime.Now);
			}
		}
	}
