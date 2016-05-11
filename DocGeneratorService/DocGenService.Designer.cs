namespace DocGeneratorService
	{
	partial class DocGeneratorServiceBase
		{
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
			{
			if(disposing && (components != null))
				{
				components.Dispose();
				}
			base.Dispose(disposing);
			}

		#region Component Designer generated code

		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
			{
			this.components = new System.ComponentModel.Container();
			this.objEventLog = new System.Diagnostics.EventLog();
			this.DocumentGenerateTimer = new System.Windows.Forms.Timer(this.components);
			this.DataRefreshTimer = new System.Windows.Forms.Timer(this.components);
			((System.ComponentModel.ISupportInitialize)(this.objEventLog)).BeginInit();
			// 
			// objEventLog
			// 
			this.objEventLog.EnableRaisingEvents = true;
			this.objEventLog.Log = "DocGenEventLog";
			// 
			// DocumentGenerateTimer
			// 
			this.DocumentGenerateTimer.Interval = 60000;
			this.DocumentGenerateTimer.Tick += new System.EventHandler(this.DocumentGenerateTimer_Tick);
			// 
			// DataRefreshTimer
			// 
			this.DataRefreshTimer.Interval = 10000;
			this.DataRefreshTimer.Tick += new System.EventHandler(this.DataRefreshTimer_Tick);
			// 
			// DocGeneratorServiceBase
			// 
			this.ServiceName = "DocGenService";
			((System.ComponentModel.ISupportInitialize)(this.objEventLog)).EndInit();

			}

		#endregion

		private System.Diagnostics.EventLog objEventLog;
		private System.Windows.Forms.Timer DocumentGenerateTimer;
		private System.Windows.Forms.Timer DataRefreshTimer;
		}
	}
