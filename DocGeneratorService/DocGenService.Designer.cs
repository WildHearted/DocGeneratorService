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
			this.objTimer = new System.Windows.Forms.Timer(this.components);
			((System.ComponentModel.ISupportInitialize)(this.objEventLog)).BeginInit();
			// 
			// objEventLog
			// 
			this.objEventLog.EnableRaisingEvents = true;
			// 
			// objTimer
			// 
			this.objTimer.Interval = 1000;
			this.objTimer.Tick += new System.EventHandler(this.objTimer_Tick);
			// 
			// DocGenService
			// 
			this.ServiceName = "DocGenService";
			((System.ComponentModel.ISupportInitialize)(this.objEventLog)).EndInit();

			}

		#endregion

		private System.Diagnostics.EventLog objEventLog;
		private System.Windows.Forms.Timer objTimer;
		}
	}
