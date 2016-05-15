namespace DocGeneratorService
	{
	partial class ProjectInstaller
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
			this.objServiceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
			this.objServiceInstaller = new System.ServiceProcess.ServiceInstaller();
			// 
			// objServiceProcessInstaller
			// 
			this.objServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
			this.objServiceProcessInstaller.Password = null;
			this.objServiceProcessInstaller.Username = null;
			// 
			// objServiceInstaller
			// 
			this.objServiceInstaller.Description = "DocGenerator Service";
			this.objServiceInstaller.DisplayName = "DocGenerator Service";
			this.objServiceInstaller.ServiceName = "DocGenService";
			// 
			// ProjectInstaller
			// 
			this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.objServiceProcessInstaller,
            this.objServiceInstaller});

			}

		#endregion

		private System.ServiceProcess.ServiceProcessInstaller objServiceProcessInstaller;
		private System.ServiceProcess.ServiceInstaller objServiceInstaller;
		}
	}