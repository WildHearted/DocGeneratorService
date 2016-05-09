using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Data.Services.Client;
using System.Runtime.InteropServices;
using System.Net;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using DocGeneratorService.SDDPServiceReference;
using DocGeneratorCore;

namespace DocGeneratorService
	{
	public partial class DocGeneratorServiceBase : System.ServiceProcess. ServiceBase
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

		public bool bBusyGenerating = false;
		public string strDocCollectionsToGenetate = String.Empty;
		public string strExceptionMessage = String.Empty;

		//[DllImport("advapi32.dll", SetLastError = true)]
		//private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);
	
		public DocGeneratorServiceBase()
			{
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " DocGeneratorServiceBase begin...");

			InitializeComponent();
			Console.WriteLine("DocGenerator - InitializeComponent completed");
			//EventLog objEventLog = new EventLog();
			//if(!EventLog.SourceExists("DocGeneratorEventSource"))
			//	{
			//	EventLog.CreateEventSource(source: "DocGeneratorEventSource", logName: "DocGenLog");
			//	}
			//objEventLog.Source = "DocGeneratorEventSource";
			//objEventLog.Log = "DocGenLog";
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " DocGeneratorServiceBase Ended...");
			}

		protected override void OnStart(string[] args)
			{
			try
				{
				Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnStart Event begin...");
				//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " +++ Starting DocGenerator Service +++");

				//System.Diagnostics.Debugger.Break();

				// Update the service state to Start Pending. 
				//ServiceStatus objServiceStatus = new ServiceStatus();
				//objServiceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
				//objServiceStatus.dwWaitHint = 30000; // 30 seconds
				//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

				//Define and set the timer's tick interval
				System.Timers.Timer objTimer = new System.Timers.Timer();
				objTimer.Interval = 60000; // timer fire every 60 seconds.

				// Inistiate the Timer's Event Ticker
				objTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.objTimer_Tick);

				// Switch the envent on
				objTimer.Enabled = true;
				objTimer.Start();

				// Update the service state to Running. 
				//objServiceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
				//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

				//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " +++ Started DocGenerator Service +++");
				}
			catch(Exception exc)
				{
				//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " +++ Exception " + exc.HResult + " occurred in OnStart service +++"
				//	+ "\n" + exc.Message + "\n" + exc.InnerException);
				Console.WriteLine(DateTime.Now.ToString("G") + " +++ Exception " + exc.HResult + " occurred in OnStart service +++"
					+ "\n" + exc.Message + "\n" + exc.InnerException);
				throw;
				}
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnStart event Ended...");
			}

		protected override void OnStop()
			{
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnStop event Begin...");
			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " --- Starting DocGenerator Service ---");
			
			// Update the service state to Stop Pending. 
			//ServiceStatus objServiceStatus = new ServiceStatus();
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
			//objServiceStatus.dwWaitHint = 300000; // 300 seconds - 5 minutes
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objTimer.Enabled = false;
			objTimer.Stop();

			// Update the service state to Stopped. 
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " --- DocGenerator Service Started ---");
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnStop event Ended...");
			}

		protected override void OnPause()
			{
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnPause event begin...");
			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === Pausing DocGenerator Service ===");

			// Update the service state to Pause Pending. 
			//ServiceStatus objServiceStatus = new ServiceStatus();
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_PAUSE_PENDING;
			//objServiceStatus.dwWaitHint = 300000; // 300 seconds - 5 minutes
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objTimer.Enabled = false;
			objTimer.Stop();

			// Update the service state to Stopped. 
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_PAUSED;
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === DocGenerator Service Paused ===");
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnPause event Ended...");
			}

		protected override void OnContinue()
			{
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnContinue event Begin...");
			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === Continuing DocGenerator Service ===");

			// Update the service state to Start Pending. 
			//ServiceStatus objServiceStatus = new ServiceStatus();
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_CONTINUE_PENDING;
			//objServiceStatus.dwWaitHint = 30000; // 30 seconds
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objTimer.Enabled = true;
			objTimer.Start();

			// Update the service state to Running. 
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === DocGenerator Service Continued ===");
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnConinue event Ended...");
			}

		protected override void OnShutdown()
			{
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnShutdown event begin...");
			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === Shutting Down DocGenerator Service ===");

			// Update the service state to Shatdown Pending. 
			//ServiceStatus objServiceStatus = new ServiceStatus();
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
			//objServiceStatus.dwWaitHint = 3000000; // 300 seconds - 5 minutes
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			objTimer.Enabled = false;
			objTimer.Stop();

			// Update the service state to Stopped. 
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);
			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === DocGenerator Service Stopped/Shutdown ===");
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnShutdown event Ended...");
			}


		private void objTimer_Tick(object sender, EventArgs e)
			{
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnTimer (tick) event begin...");
			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " +++ DocGenerator Service event fired +++");

			if(!bBusyGenerating)
				{
				try
					{
					bBusyGenerating = true;
					strDocCollectionsToGenetate = String.Empty;
					string strExceptionMessage = string.Empty;

					//Construct the SharePoint Client Context
					DesignAndDeliveryPortfolioDataContext objSDDPdatacontext = new DesignAndDeliveryPortfolioDataContext(new Uri(Properties.Resources.SharePointSiteURL + Properties.Resources.SharePointRESTuri));

					objSDDPdatacontext.Credentials = new NetworkCredential(
						userName: Properties.Resources.DocGenerator_AccountName,
						password: Properties.Resources.DocGenerator_Account_Password,
						domain: Properties.Resources.DocGenerator_AccountDomain);

					objSDDPdatacontext.MergeOption = MergeOption.NoTracking;

					var dsDocCollections = from dsDocumentCollection in objSDDPdatacontext.DocumentCollectionLibrary
									where dsDocumentCollection.GenerateActionValue != null
									&& dsDocumentCollection.GenerateActionValue != "Save but don't generate the documents yet"
									&& dsDocumentCollection.GenerationStatus != enumGenerationStatus.Completed.ToString()
									&& dsDocumentCollection.GenerationStatus != enumGenerationStatus.Failed.ToString()
									&& dsDocumentCollection.GenerationStatus != enumGenerationStatus.Generating.ToString()
									orderby dsDocumentCollection.Modified select dsDocumentCollection;

					foreach(var objDocCollectionToGenerate in dsDocCollections)
						{
						strDocCollectionsToGenetate += objDocCollectionToGenerate.Id + ",";
						}

					if(strDocCollectionsToGenetate != String.Empty)
						{
						// Invoke the DocGeneratorCore's MainController object MainProcess method
						MainController objMainController = new MainController();

						objMainController.MainProcess();
						bBusyGenerating = false;
						}
					}
				catch(DataServiceClientException exc)
					{
					Console.Beep(2500, 750);
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult + "\nStatusCode: " + exc.StatusCode
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Console.WriteLine(strExceptionMessage);
					throw new GeneralException(strExceptionMessage);
					}
				catch(DataServiceQueryException exc)
					{
					Console.Beep(2500, 750);
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Console.WriteLine(strExceptionMessage);
					throw new GeneralException(strExceptionMessage);
					}
				catch(DataServiceRequestException exc)
					{
					Console.Beep(2500, 750);
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Console.WriteLine(strExceptionMessage);
					throw new GeneralException(strExceptionMessage);
					}
				catch(DataServiceTransportException exc)
					{
					Console.Beep(2500, 750);
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Console.WriteLine(strExceptionMessage);
					throw new GeneralException(strExceptionMessage);
					}
				catch(Exception exc)
					{
					Console.Beep(2500, 750);

					if(exc.HResult == -2146330330)
						{
						strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
						}
					else if(exc.HResult == -2146233033)
						{
						strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
						}
					else
						{
						strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
						};
					}
				}

			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " +++ Event Ended...");
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnTimer (tick) event Ended...");
			}


		}
	}
