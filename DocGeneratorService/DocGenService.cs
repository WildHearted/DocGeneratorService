using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Services.Client;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Runtime.InteropServices;
using System.Net;
using System.Threading;
using Microsoft.SharePoint.Client;
using DocGeneratorCore;
using DocGeneratorService.SDDPServiceReference;


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

		private Object objThreadLock = new Object();
		private DesignAndDeliveryPortfolioDataContext SDDPdatacontext = new DesignAndDeliveryPortfolioDataContext(
			new Uri(Properties.Resources.SharePointSiteURL + Properties.Resources.SharePointRESTuri));
		private DateTime dtDataRefreshed = new DateTime(2000,1,1,0,0,0);
		private static bool bCompleteDataSetReady = false;
		private CompleteDataSet objCompleteDataSet = new CompleteDataSet();
		//private static Barrier barrierDataSetRefesh = new Barrier(
		//	participantCount: 5, 
		//	postPhaseAction: (bar) =>
		//	{
		//		if()
		//	});

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
				//System.Timers.Timer objTimer = new System.Timers.Timer();
				//objTimer.Interval = 60000; // timer fire every 60 seconds.
				// Initiate the Timer's Event Ticker
				//objTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.DocumentGenerateTimer_Tick);
				// Switch the envent on
				//objTimer.Enabled = true;
				//objTimer.Start();

				// Configure the timer for the refreshing the DataSet
				this.DataRefreshTimer.Interval = 10000; // Timer to fire every 10 seconds
				this.DataRefreshTimer.Enabled = true;
				this.DataRefreshTimer.Start();

				// Configure the timer for the Generation of Documents
				this.DocumentGenerateTimer.Interval = 60000; // Timer to fire every 60 seconds
				this.DocumentGenerateTimer.Enabled = true;
				this.DocumentGenerateTimer.Start();

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

			DocumentGenerateTimer.Enabled = false;
			DocumentGenerateTimer.Stop();

			DataRefreshTimer.Enabled = false;
			DataRefreshTimer.Stop();

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

			DocumentGenerateTimer.Enabled = false;
			DocumentGenerateTimer.Stop();

			DataRefreshTimer.Enabled = false;
			DataRefreshTimer.Stop();

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

			DocumentGenerateTimer.Enabled = true;
			DocumentGenerateTimer.Start();

			DataRefreshTimer.Enabled = false;
			DataRefreshTimer.Stop();
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

			DocumentGenerateTimer.Enabled = false;
			DocumentGenerateTimer.Stop();

			DataRefreshTimer.Enabled = false;
			DataRefreshTimer.Stop();
			// Update the service state to Stopped. 
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);
			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === DocGenerator Service Stopped/Shutdown ===");
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnShutdown event Ended...");
			}


		private void DataRefreshTimer_Tick(object sender, EventArgs e)
			{
			//this.SDDPdatacontext.Credentials = new NetworkCredential(
			//			userName: Properties.Resources.DocGenerator_AccountName,
			//			password: Properties.Resources.DocGenerator_Account_Password,
			//			domain: Properties.Resources.DocGenerator_AccountDomain);
			//SDDPdatacontext.MergeOption = MergeOption.NoTracking;

			//Thread objThread1 = new Thread(() => objCompleteDataSet.PopulateBaseObjects());


			}

		private void DocumentGenerateTimer_Tick(object sender, EventArgs e)
			{
			// Protect this code that only a single thread at a time can generate documents.
			lock (objThreadLock)
				{
				String EmailBodyText = String.Empty;
				string strExceptionMessage = String.Empty;
				bool bSuccessfulSentEmail = false;
				List<DocumentCollection> listDocumentCollections = new List<DocumentCollection>();
	
				try
					{
					//Construct the SharePoint Client Context
					//DesignAndDeliveryPortfolioDataContext objSDDPdatacontext = new DesignAndDeliveryPortfolioDataContext(
					//	new Uri(Properties.Resources.SharePointSiteURL + Properties.Resources.SharePointRESTuri));

					SDDPdatacontext.Credentials = new NetworkCredential(
						userName: Properties.Resources.DocGenerator_AccountName,
						password: Properties.Resources.DocGenerator_Account_Password,
						domain: Properties.Resources.DocGenerator_AccountDomain);

					SDDPdatacontext.MergeOption = MergeOption.NoTracking;

					var dsDocCollections = from dsDocumentCollection in SDDPdatacontext.DocumentCollectionLibrary
									   where dsDocumentCollection.GenerateActionValue != null
										&& dsDocumentCollection.GenerateActionValue != "Save but don't generate the documents yet"
										&& (dsDocumentCollection.GenerationStatus == enumGenerationStatus.Pending.ToString()
											|| dsDocumentCollection.GenerationStatus == null)
										orderby dsDocumentCollection.Modified
										select dsDocumentCollection;

					foreach(var recDocCollectionToGenerate in dsDocCollections)
						{
						// Create a DocumentCollection instance and populate the basic attributes.
						DocumentCollection objDocumentCollection = new DocumentCollection();
						objDocumentCollection.ID = recDocCollectionToGenerate.Id;
						if(recDocCollectionToGenerate.Title == null)
							objDocumentCollection.Title = "Collection Title for entry " + recDocCollectionToGenerate.Id;
						else
							objDocumentCollection.Title = recDocCollectionToGenerate.Title;
						objDocumentCollection.DetailComplete = false;
						// Add the DocumentCollection object to the listDocumentCollection
						listDocumentCollections.Add(objDocumentCollection);
						}

					// Check if there are any Document Collections to generate
					if(listDocumentCollections.Count > 0)
						{
						foreach(DocumentCollection entryDocumentCollection in listDocumentCollections)
							{// Invoke the DocGeneratorCore's MainController object MainProcess method
							MainController objMainController = new MainController();
							objMainController.DocumentCollectionsToGenerate = listDocumentCollections;
							objMainController.MainProcess();
							}
						}
					}
				catch(DataServiceClientException exc)
					{
					strExceptionMessage = "*** Exception ERROR ***: DocGeneratorServer cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult + "\nStatusCode: " + exc.StatusCode
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
					bSuccessfulSentEmail = eMail.SendEmail(
						parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
						parSubject: "Error occurred in DocGenerator Server module.)",
						parBody: EmailBodyText,
						parSendBcc: false);
					}
				catch(DataServiceQueryException exc)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
					bSuccessfulSentEmail = eMail.SendEmail(
						parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
						parSubject: "Error occurred in DocGenerator Server module.)",
						parBody: EmailBodyText,
						parSendBcc: false);
					}
				catch(DataServiceRequestException exc)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
					bSuccessfulSentEmail = eMail.SendEmail(
						parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
						parSubject: "Error occurred in DocGenerator Server module.)",
						parBody: EmailBodyText,
						parSendBcc: false);
					}
				catch(DataServiceTransportException exc)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
					bSuccessfulSentEmail = eMail.SendEmail(
						parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
						parSubject: "Error occurred in DocGenerator Server module.)",
						parBody: EmailBodyText,
						parSendBcc: false);
					}
				catch(Exception exc)
					{
					if(exc.HResult == -2146330330)
						{
						strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
						EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
						bSuccessfulSentEmail = eMail.SendEmail(
							parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
							parSubject: "Error occurred in DocGenerator Server module.)",
							parBody: EmailBodyText,
							parSendBcc: false);
						}
					else if(exc.HResult == -2146233033)
						{
						strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
						EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
						bSuccessfulSentEmail = eMail.SendEmail(
							parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
							parSubject: "Error occurred in DocGenerator Server module.)",
							parBody: EmailBodyText,
							parSendBcc: false);
						}
					else
						{
						strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
						EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
						bSuccessfulSentEmail = eMail.SendEmail(
							parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
							parSubject: "Error occurred in DocGenerator Server module.)",
							parBody: EmailBodyText,
							parSendBcc: false);
						}
					} //Catch
				} // End Lock...
			}
		}
	}
