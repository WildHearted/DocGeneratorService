using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Services.Client;
using System.Data.Linq;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;
using DocGeneratorCore;
using Microsoft.SharePoint.Client;


namespace DocGeneratorService
	{
	public partial class DocGeneratorServiceBase:System.ServiceProcess.ServiceBase
		{
		//- the *shutdownEvent** controls the Shutting down of th service.
		private ManualResetEvent shutdownEvent = new ManualResetEvent(initialState: false);
		private NetworkCredential credentialsDocGenerator = new NetworkCredential();
		//- The docGeneratorThread 
		private Thread docGeneratorThread;
		private CompleteDataSet completeDataSet = new CompleteDataSet();

		public DocGeneratorServiceBase()
			{
			//- uncomment the following line to debug the service
			//Debugger.Break();

			InitializeComponent();

			//EventLog.Log = "DocGenerator";

			}

		protected override void OnStart(string[] args)
			{
			//- Establish the SDDPdatacontect that will be used to access data on SharePoint

			//this.completeDataSet.SDDPdatacontext = new DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext(
			//	new Uri(Properties.Resources.SharePointSiteURL + Properties.Resources.SharePointRESTuri));

			//this.completeDataSet.SDDPdatacontext.Credentials = new NetworkCredential(
			//- Set the Credetials that the DocGenerator application use during processing
			credentialsDocGenerator = new NetworkCredential(				
					userName: Properties.Resources.DocGenerator_AccountName,
					password: Properties.Resources.DocGenerator_Account_Password,
					domain: Properties.Resources.DocGenerator_AccountDomain);
			//this.completeDataSet.SDDPdatacontext.MergeOption = MergeOption.NoTracking;



			//- Create the Thread which will execute service
			docGeneratorThread = new Thread(ServiceControlFunction);
			docGeneratorThread.Name = "Main Service Thread";
			docGeneratorThread.IsBackground = true;
			docGeneratorThread.Start();

			//EventLog.WriteEntry("DocGenerator Service started successfully", EventLogEntryType.Information);

			}

		private void ServiceControlFunction()
			{
			string EmailBodyText = String.Empty;
			string strExceptionMessage = String.Empty;
			bool bSuccessfulSentEmail = false;

			// Notify Technical Support that the DocGenerator Service  started.
			TechnicalSupportModel emailModel = new TechnicalSupportModel();
			emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
			emailModel.Classification = enumMessageClassification.Information;
			emailModel.MessageHeading = "For your information...";
			emailModel.Instruction = "This is a just to inform you that...";
			emailModel.MessageLines = new List<string>();
			emailModel.MessageLines.Add("The DocGenerator Service started at " + DateTime.Now.ToString());
			emailModel.MessageLines.Add("No need to worry...");
			//- Declare the Email object and assign the above defined message to the relevant property
			eMail objTechnicalEmail = new eMail();
			objTechnicalEmail.TechnicalEmailModel = emailModel;
			//- Compile the HTML email message
			if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
				{
				bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
					parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
					parSubject: "DocGenerator: Service Started: " + DateTime.UtcNow.ToString(),
					parSendBcc: false);
				}

			// Check if a network connection is available
			if(!IsNetworkAvailable())
				{
				emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
				emailModel.Classification = enumMessageClassification.Warning;
				emailModel.MessageHeading = "For your information...";
				emailModel.Instruction = "On service startup, the DocGenerator didn't detect a Network connection...";
				emailModel.MessageLines = new List<string>();
				emailModel.MessageLines.Add("At " + DateTime.Now.ToString() + "the DocGenerator could not detect a network connection.");
				emailModel.MessageLines.Add("Please watch out for any other connectivity failures.");
				emailModel.MessageLines.Add("If more connectivity failures occur, please investigate the connectivity status of the DocGenerator server.");
				// Declare the Email object and assign the above defined message to the relevant property
				objTechnicalEmail = new eMail();
				objTechnicalEmail.TechnicalEmailModel = emailModel;
				//- Compile the HTML email message
				if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
					{
					bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
						parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
						parSubject: "DocGenerator: Connectivity Warning: " + DateTime.UtcNow.ToString(),
						parSendBcc: false);
					}
				}
			// --------------------------------------------------------------------------------------------------------
			//+ Infinite loop, which will only stop when the **shutdownEvent** become true
			// --------------------------------------------------------------------------------------------------------
			while(!shutdownEvent.WaitOne(0))
				{
				EmailBodyText = String.Empty;
				strExceptionMessage = String.Empty;
				bSuccessfulSentEmail = false;
				List<DocumentCollection> listDocumentCollections;

				// Check if a network connection is available before checking for Entries to generate...
				if(!IsNetworkAvailable())
					{ // - if NOT, send a Technical Support e-mail
					emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
					emailModel.Classification = enumMessageClassification.Error;
					emailModel.MessageHeading = "Connectivity Error";
					emailModel.Instruction = "The DocGenerator didn't detect a Network connection...";
					emailModel.MessageLines = new List<string>();
					emailModel.MessageLines.Add("At " + DateTime.Now.ToString() + "the DocGenerator could not detect a network connection.");
					emailModel.MessageLines.Add("Please investigate the connectivity status of the DocGenerator server.");
					// Declare the Email object and assign the above defined message to the relevant property
					objTechnicalEmail = new eMail();
					objTechnicalEmail.TechnicalEmailModel = emailModel;
					//- Compile the HTML email message
					if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
						{
						bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
							parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
							parSubject: "DocGenerator: Connectivity ERROR: " + DateTime.UtcNow.ToString(),
							parSendBcc: false);
						}
					}
				else
					{//- if connectivity is ok, continue to process...
					 //- Check if the SharePoint environment can be reached
					HttpWebRequest objHTTPwebRequest = WebRequest.Create(
						requestUriString: Properties.Resources.SharePointSiteURL) as HttpWebRequest;
					objHTTPwebRequest.Credentials = credentialsDocGenerator;
					objHTTPwebRequest.Timeout = 15000;
					HttpWebResponse objHTTPwebResponse;
					try
						{
						objHTTPwebResponse = objHTTPwebRequest.GetResponse() as HttpWebResponse;
						}

					catch (NullReferenceException)
						{
						objHTTPwebResponse = null;
						}

					catch (ArgumentNullException)
						{
						objHTTPwebResponse = null;
						}

					catch(WebException exc)
						{
						if(exc.Response is HttpWebResponse)
							{
							objHTTPwebResponse = exc.Response as HttpWebResponse;
							}
						else
							{
							objHTTPwebResponse = null;
							}
						}
					if(objHTTPwebResponse == null)
						{//- if valid response received, send a Technical Support e-mail
						emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
						emailModel.Classification = enumMessageClassification.Error;
						emailModel.MessageHeading = "DocGenerator: Unable to contact SharePoint";
						emailModel.Instruction = "The DocGenerator couldn't reach the SharePoint at " + Properties.Resources.SharePointSiteURL;
						emailModel.MessageLines = new List<string>();
						emailModel.MessageLines.Add("At " + DateTime.Now.ToString() + "the DocGenerator could not reach the SDDP SharePoint environment.");
						emailModel.MessageLines.Add("Please investigate the connectivity status of the DocGenerator server.");
						// Declare the Email object and assign the above defined message to the relevant property
						objTechnicalEmail = new eMail();
						objTechnicalEmail.TechnicalEmailModel = emailModel;
						//- Compile the HTML email message
						if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
							{
							bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
								parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
								parSubject: "DocGenerator: Unable to contact SharePoint @ " + DateTime.UtcNow.ToString(),
								parSendBcc: false);
							}
						}
					else
						{ //- The SharePoint Server was SUCCESSFULLY contacted
						//- Establish a new SDDPdatacontect that will be used to access data on SharePoint
						try
							{
							//- Establish the DataContext to read the data from SharePoint
							DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext datacontectSDDP = new DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext(
								new Uri(Properties.Resources.SharePointSiteURL + Properties.Resources.SharePointRESTuri));

							datacontectSDDP.Credentials = credentialsDocGenerator;
							datacontectSDDP.MergeOption = MergeOption.NoTracking;

							//this.completeDataSet.SDDPdatacontext = new DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext(
							//	new Uri(Properties.Resources.SharePointSiteURL + Properties.Resources.SharePointRESTuri));
							//this.completeDataSet.SDDPdatacontext.Credentials = credentialsDocGenerator;
							//this.completeDataSet.SDDPdatacontext.MergeOption = MergeOption.NoTracking;

							//- Initialise a List of DocumentCollections - the list will contain all the DocumentCollection objects that need to be generated in this cycle.
							listDocumentCollections = new List<DocumentCollection>();
						
							//- Get all the DocumentCollections that require Generation
							var dsDocCollections =
								from dsDocumentCollection in datacontectSDDP.DocumentCollectionLibrary 
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
									objMainController.MainProcess(parDataSet: ref completeDataSet);
									}
								}
							else  
								{//- there are no Document collections to generate, therefore the thread will sleep for 60 seconds
								this.completeDataSet.SDDPdatacontext = null;
								//- Pause the thread for 60 seconds...
								System.Threading.Thread.Sleep(60000);
								}
							}

						catch(DataServiceClientException exc)
							{
							emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
							emailModel.Classification = enumMessageClassification.Warning;
							emailModel.MessageHeading = "Warning error occurred in DocGenerator Server module.";
							emailModel.Instruction = "DocGenerator cannot access the SharePoint site: " + Properties.Resources.SharePointSiteURL;
							emailModel.MessageLines = new List<string>();
							emailModel.MessageLines.Add("The following DataServiceClientException occurred at " + DateTime.Now.ToString());
							emailModel.MessageLines.Add("HResult: " + exc.HResult);
							emailModel.MessageLines.Add("Error Message: " + exc.Message);
							emailModel.MessageLines.Add("InnerException: " + exc.InnerException);
							emailModel.MessageLines.Add("Please check that the computer/server is connected to the Domain network ");
							// Declare the Email object and assign the above defined message to the relevant property
							objTechnicalEmail = new eMail();
							objTechnicalEmail.TechnicalEmailModel = emailModel;
							//- Compile the HTML email message
							if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
								{
								bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
									parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
									parSubject: "DocGenerator: Warning occurred in Service Module at: " + DateTime.UtcNow.ToString(),
									parSendBcc: false);
								}
							}

						catch(DataServiceQueryException exc)
							{
							emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
							emailModel.Classification = enumMessageClassification.Warning;
							emailModel.MessageHeading = "Warning error occurred in DocGenerator Server module.";
							emailModel.Instruction = "DocGenerator cannot access the SharePoint site: " + Properties.Resources.SharePointSiteURL;
							emailModel.MessageLines = new List<string>();
							emailModel.MessageLines.Add("The following DataServiceQueryException occurred at " + DateTime.Now.ToString());
							emailModel.MessageLines.Add("HResult: " + exc.HResult);
							emailModel.MessageLines.Add("Error Message: " + exc.Message);
							emailModel.MessageLines.Add("InnerException: " + exc.InnerException);
							emailModel.MessageLines.Add("Please check that the computer/server is connected to the Domain network ");
							// Declare the Email object and assign the above defined message to the relevant property
							objTechnicalEmail = new eMail();
							objTechnicalEmail.TechnicalEmailModel = emailModel;
							//- Compile the HTML email message
							if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
								{
								bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
									parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
									parSubject: "DocGenerator: Warning occurred in Service Module at: " + DateTime.UtcNow.ToString(),
									parSendBcc: false);
								}
							}

						catch(DataServiceRequestException exc)
							{
							emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
							emailModel.Classification = enumMessageClassification.Warning;
							emailModel.MessageHeading = "Warning error occurred in DocGenerator Server module.";
							emailModel.Instruction = "DocGenerator cannot access the SharePoint site: " + Properties.Resources.SharePointSiteURL;
							emailModel.MessageLines = new List<string>();
							emailModel.MessageLines.Add("The following DataServiceRequestException occurred at " + DateTime.Now.ToString());
							emailModel.MessageLines.Add("HResult: " + exc.HResult);
							emailModel.MessageLines.Add("Error Message: " + exc.Message);
							emailModel.MessageLines.Add("InnerException: " + exc.InnerException);
							emailModel.MessageLines.Add("Please check that the computer/server is connected to the Domain network ");
							// Declare the Email object and assign the above defined message to the relevant property
							objTechnicalEmail = new eMail();
							objTechnicalEmail.TechnicalEmailModel = emailModel;
							//- Compile the HTML email message
							if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
								{
								bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
									parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
									parSubject: "DocGenerator: Warning occurred in Service Module at: " + DateTime.UtcNow.ToString(),
									parSendBcc: false);
								}
							}

						catch(DataServiceTransportException exc)
							{
							emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
							emailModel.Classification = enumMessageClassification.Warning;
							emailModel.MessageHeading = "Warning error occurred in DocGenerator Server module.";
							emailModel.Instruction = "DocGenerator cannot access the SharePoint site: " + Properties.Resources.SharePointSiteURL;
							emailModel.MessageLines = new List<string>();
							emailModel.MessageLines.Add("The following DataServiceTransportException occurred at " + DateTime.Now.ToString());
							emailModel.MessageLines.Add("HResult: " + exc.HResult);
							emailModel.MessageLines.Add("Error Message: " + exc.Message);
							emailModel.MessageLines.Add("InnerException: " + exc.InnerException);
							emailModel.MessageLines.Add("Please check that the computer/server is connected to the Domain network ");
							// Declare the Email object and assign the above defined message to the relevant property
							objTechnicalEmail = new eMail();
							objTechnicalEmail.TechnicalEmailModel = emailModel;
							//- Compile the HTML email message
							if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
								{
								bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
									parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
									parSubject: "DocGenerator: Warning occurred in Service Module at: " + DateTime.UtcNow.ToString(),
									parSendBcc: false);
								}
							}

						catch(Exception exc)
							{
							if(exc.HResult == -2146330330)
								{
								emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
								emailModel.Classification = enumMessageClassification.Warning;
								emailModel.MessageHeading = "Warning error occurred in DocGenerator Server module.";
								emailModel.Instruction = "DocGenerator cannot access the SharePoint site: " + Properties.Resources.SharePointSiteURL;
								emailModel.MessageLines = new List<string>();
								emailModel.MessageLines.Add("The following Undefined Exception occurred at " + DateTime.Now.ToString());
								emailModel.MessageLines.Add("HResult: " + exc.HResult);
								emailModel.MessageLines.Add("Error Message: " + exc.Message);
								emailModel.MessageLines.Add("InnerException: " + exc.InnerException);
								emailModel.MessageLines.Add("Please check that the computer/server is connected to the Domain network ");
								// Declare the Email object and assign the above defined message to the relevant property
								objTechnicalEmail = new eMail();
								objTechnicalEmail.TechnicalEmailModel = emailModel;
								//- Compile the HTML email message
								if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
									{
									bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
										parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
										parSubject: "DocGenerator: Warning occurred in Service Module at: " + DateTime.UtcNow.ToString(),
										parSendBcc: false);
									}
								}
							else if(exc.HResult == -2146233033)
								{
								emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
								emailModel.Classification = enumMessageClassification.Warning;
								emailModel.MessageHeading = "Warning error occurred in DocGenerator Server module.";
								emailModel.Instruction = "DocGenerator cannot access the SharePoint site: " + Properties.Resources.SharePointSiteURL;
								emailModel.MessageLines = new List<string>();
								emailModel.MessageLines.Add("The following Undefined Exception occurred at " + DateTime.Now.ToString());
								emailModel.MessageLines.Add("HResult: " + exc.HResult);
								emailModel.MessageLines.Add("Error Message: " + exc.Message);
								emailModel.MessageLines.Add("InnerException: " + exc.InnerException);
								emailModel.MessageLines.Add("Please check that the computer/server is connected to the Domain network ");
								// Declare the Email object and assign the above defined message to the relevant property
								objTechnicalEmail = new eMail();
								objTechnicalEmail.TechnicalEmailModel = emailModel;
								//- Compile the HTML email message
								if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
									{
									bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
										parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
										parSubject: "DocGenerator: Warning occurred in Service Module at: " + DateTime.UtcNow.ToString(),
										parSendBcc: false);
									}
								}
							else
								{
								emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
								emailModel.Classification = enumMessageClassification.Error;
								emailModel.MessageHeading = "Unexpected error occurred in DocGenerator Server module.";
								emailModel.Instruction = "DocGenerator cannot access the SharePoint site: " + Properties.Resources.SharePointSiteURL;
								emailModel.MessageLines = new List<string>();
								emailModel.MessageLines.Add("The following Undefined Exception occurred at " + DateTime.Now.ToString());
								emailModel.MessageLines.Add("HResult: " + exc.HResult);
								emailModel.MessageLines.Add("Error Message: " + exc.Message);
								emailModel.MessageLines.Add("InnerException: " + exc.InnerException);
								emailModel.MessageLines.Add("Please check that the computer/server is connected to the Domain network ");
								// Declare the Email object and assign the above defined message to the relevant property
								objTechnicalEmail = new eMail();
								objTechnicalEmail.TechnicalEmailModel = emailModel;
								//- Compile the HTML email message
								if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
									{
									bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
										parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
										parSubject: "DocGenerator: Warning occurred in Service Module at: " + DateTime.UtcNow.ToString(),
										parSendBcc: false);
									}
								}
							}
						}
					}
				}
			}



		protected override void OnStop()
			{
			//- signal the **Stop Event**
			//- Send an e-mail to the Technical support team when the service is stopped
			TechnicalSupportModel emailModel = new TechnicalSupportModel();
			emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
			emailModel.Classification = enumMessageClassification.Warning;
			emailModel.MessageHeading = "Warning: the DocGenerator Server Service stopperd.";
			emailModel.Instruction = "This is a Warning message to inform you that...";
			emailModel.MessageLines = new List<string>();
			emailModel.MessageLines.Add("The DocGenerator Service was stopped at " + DateTime.UtcNow.ToString());
			emailModel.MessageLines.Add("Please watch your e-mail and investigate of the Service does'n restart in 5 to 10 minutes.");
			// Declare the Email object and assign the above defined message to the relevant property
			eMail objTechnicalEmail = new eMail();
			objTechnicalEmail.TechnicalEmailModel = emailModel;
			//- Compile the HTML email message
			if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
				{
				bool bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
					parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
					parSubject: "DocGenerator: Warning occurred in Service Module at: " + DateTime.UtcNow.ToString(),
					parSendBcc: false);
				}
			
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
			//- signal the **shutdownEvent**
			//- Send an e-mail to the Technical support team when the service is stopped
			TechnicalSupportModel emailModel = new TechnicalSupportModel();
			emailModel.EmailAddress = Properties.Resources.EmailAddress_TechnicalSupport;
			emailModel.Classification = enumMessageClassification.Warning;
			emailModel.MessageHeading = "Warning: the DocGenerator Server Service stopperd.";
			emailModel.Instruction = "This is a Warning message to inform you that...";
			emailModel.MessageLines = new List<string>();
			emailModel.MessageLines.Add("The DocGenerator Service was stopped at " + DateTime.UtcNow.ToString());
			emailModel.MessageLines.Add("Please watch your e-mail and investigate of the Service does'n restart in 5 to 10 minutes.");
			// Declare the Email object and assign the above defined message to the relevant property
			eMail objTechnicalEmail = new eMail();
			objTechnicalEmail.TechnicalEmailModel = emailModel;
			//- Compile the HTML email message
			if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
				{
				bool bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
					parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
					parSubject: "DocGenerator: Warning occurred in Service Module at: " + DateTime.UtcNow.ToString(),
					parSendBcc: false);
				}

			shutdownEvent.Set();
			if(docGeneratorThread.Join(5000))
				{
				docGeneratorThread.Abort();
				}

			//EventLog.WriteEntry("DocGenerator Service SHUTDOWN successfully", EventLogEntryType.Information);

			}

		/// <summary>
		/// Indicates whether any network connection is available
		/// Filter connections below a specified speed, as well as virtual network cards.
		/// </summary>
		/// <returns>
		///     <c>true</c> if a network connection is available; otherwise, <c>false</c>.
		/// </returns>
		public static bool IsNetworkAvailable()
			{
			return IsNetworkAvailable(10000000);
			}

		/// <summary>
		/// Indicates whether any network connection is available.
		/// Filter connections below a specified speed, as well as virtual network cards.
		/// </summary>
		/// <param name="parMinimumSpeed">The minimum speed required. Passing 0 will not filter connection using speed.</param>
		/// <returns>
		///     <c>true</c> if a network connection is available; otherwise, <c>false</c>.
		/// </returns>
		public static bool IsNetworkAvailable(long parMinimumSpeed)
			{
			if(!NetworkInterface.GetIsNetworkAvailable())
				return false;

			foreach(NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
				{
				// discard because of standard reasons
				if((ni.OperationalStatus != OperationalStatus.Up) ||
				    (ni.NetworkInterfaceType == NetworkInterfaceType.Loopback) ||
				    (ni.NetworkInterfaceType == NetworkInterfaceType.Tunnel))
					continue;

				// this allow to filter modems, serial, etc.
				// I use 10000000 as a minimum speed for most cases
				if(ni.Speed < parMinimumSpeed)
					continue;

				// discard virtual cards (virtual box, virtual pc, etc.)
				if((ni.Description.IndexOf("virtual", StringComparison.OrdinalIgnoreCase) >= 0) ||
				    (ni.Name.IndexOf("virtual", StringComparison.OrdinalIgnoreCase) >= 0))
					continue;

				// discard "Microsoft Loopback Adapter", it will not show as NetworkInterfaceType.Loopback but as Ethernet Card.
				if(ni.Description.Equals("Microsoft Loopback Adapter", StringComparison.OrdinalIgnoreCase))
					continue;

				return true;
				}
			return false;
			}

		}
	}
