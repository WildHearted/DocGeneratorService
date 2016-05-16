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
//using DocGeneratorService.SDDPServiceReference;


namespace DocGeneratorService
	{
	public partial class DocGeneratorServiceBase:System.ServiceProcess.ServiceBase
		{
		//- the *shutdownEvent** controls the Shutting down of th service.
		private ManualResetEvent shutdownEvent = new ManualResetEvent(initialState: false);
		//- The docGeneratorThread 
		private Thread docGeneratorThread;
		private CompleteDataSet comompleteDataSet = new CompleteDataSet();

		public DocGeneratorServiceBase()
			{

			//- uncomment the following line to debug the service
			//Debugger.Break();

			InitializeComponent();

			//
			//EventLog.Log = "DocGenerator";

			}

		protected override void OnStart(string[] args)
			{

			//- Establish the SDDPdatacontect that will be used to access data on SharePoint

			this.comompleteDataSet.SDDPdatacontext = new DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext(
				new Uri(Properties.Resources.SharePointSiteURL + Properties.Resources.SharePointRESTuri));

			this.comompleteDataSet.SDDPdatacontext.Credentials = new NetworkCredential(
					userName: Properties.Resources.DocGenerator_AccountName,
					password: Properties.Resources.DocGenerator_Account_Password,
					domain: Properties.Resources.DocGenerator_AccountDomain);
			this.comompleteDataSet.SDDPdatacontext.MergeOption = MergeOption.NoTracking;

			//- Create the Thread which will execute service
			docGeneratorThread = new Thread(ServiceControlFunction);
			docGeneratorThread.Name = "Main Service Thread";
			docGeneratorThread.IsBackground = true;
			docGeneratorThread.Start();

			//EventLog.WriteEntry("DocGenerator Service started successfully", EventLogEntryType.Information);

			}

		private void ServiceControlFunction()
			{
			eMail.SendEmail(
				parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
				parSubject: "DocGenerator: Service Started: " + DateTime.UtcNow.ToString(),
				parBody: "Hi there\nJust a notification to let you know that the DocGenerator has Started again.",
				parSendBcc: false);

			//- Have an infinite loop, which will only stop when the **shutdownEvent** become true
			while(!shutdownEvent.WaitOne(0))
				{

				String EmailBodyText = String.Empty;
				string strExceptionMessage = String.Empty;
				bool bSuccessfulSentEmail = false;
				//- Initialise a List of DocumentCollections - the list will contain all the DocumentCollection objects that nedds to be generated in this cycle.
				List<DocumentCollection> listDocumentCollections = new List<DocumentCollection>();
	
				try
					{
					//- Get all the DocumentCollections that require Generation
					var dsDocCollections = 
						from dsDocumentCollection in this.comompleteDataSet.SDDPdatacontext.DocumentCollectionLibrary
						where dsDocumentCollection.GenerateActionValue != null
						&& dsDocumentCollection.GenerateActionValue != "Save but don't generate the documents yet"
						&& (dsDocumentCollection.GenerationStatus == enumGenerationStatus.Pending.ToString()
						|| dsDocumentCollection.GenerationStatus == null)
						orderby dsDocumentCollection.Modified
						select dsDocumentCollection;

					foreach(var recDocCollectionToGenerate in dsDocCollections)
						{
						//- Create a DocumentCollection instance and populate the basic attributes.
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

					//- If any DocumentCollections were found, - Generate each one...
					if(listDocumentCollections.Count > 0)
						{
						foreach(DocumentCollection entryDocumentCollection in listDocumentCollections)
							{// Invoke the DocGeneratorCore's MainController object MainProcess method
							MainController objMainController = new MainController();
							objMainController.DocumentCollectionsToGenerate = listDocumentCollections;
							objMainController.MainProcess(parDataSet: ref comompleteDataSet);
							}
						}
					else  //- there are no Document collections to generate, therefore the thread will sleep for 60 seconds
						{
						//- Pause the thread for 60 seconds...
						System.Threading.Thread.Sleep(60000);
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
					}
				}
			}



		protected override void OnStop()
			{
			//- signal the **shutdownEvent
			//- Send an e-mail to the Technical support team when the service is stopped
			eMail.SendEmail(
				parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
				parSubject: "DocGenerator: Service stopped: " + DateTime.UtcNow.ToString(),
				parBody: "Hi there\nJust a notification to let you know that the DocGenerator was stopped.",
				parSendBcc: false);

			shutdownEvent.Set();
			if(docGeneratorThread.Join(5000))
				{
				docGeneratorThread.Abort();
				}

			//EventLog.WriteEntry("DocGenerator Service stopped successfully", EventLogEntryType.Information);
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


			// Update the service state to Stopped. 
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_PAUSED;
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === DocGenerator Service Paused ===");
			Console.WriteLine("\t + " + DateTime.Now.ToString("G") + " OnPause event Ended...");
			}

		protected override void OnContinue()
			{
			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === Continuing DocGenerator Service ===");

			// Update the service state to Start Pending. 
			//ServiceStatus objServiceStatus = new ServiceStatus();
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_CONTINUE_PENDING;
			//objServiceStatus.dwWaitHint = 30000; // 30 seconds
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);


			// Update the service state to Running. 
			//objServiceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
			//SetServiceStatus(this.ServiceHandle, ref objServiceStatus);

			//objEventLog.WriteEntry(DateTime.Now.ToString("G") + " === DocGenerator Service Continued ===");

			}

		protected override void OnShutdown()
			{
			//- signal the **shutdownEvent
			shutdownEvent.Set();
			if(docGeneratorThread.Join(5000))
				{
				docGeneratorThread.Abort();
				}

			//EventLog.WriteEntry("DocGenerator Service SHUTDOWN successfully", EventLogEntryType.Information);

			}
		}
	}
