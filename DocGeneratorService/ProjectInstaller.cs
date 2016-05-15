using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace DocGeneratorService
	{
	[RunInstaller(true)]
	public partial class ProjectInstaller:System.Configuration.Install.Installer
		{
		public ProjectInstaller()
			{
			InitializeComponent();

			EventLogInstaller eventLogInstaller = FindInstaller(this.Installers);
			if(eventLogInstaller != null)
				{
				eventLogInstaller.Log = "DocGenerator";
				}
			}

		private EventLogInstaller FindInstaller(InstallerCollection installers)
			{
			foreach(Installer installerEntry in installers)
				{
				if(installerEntry is EventLogInstaller)
					{
					return (EventLogInstaller)installerEntry;
					}

				EventLogInstaller eventlogInstaller = FindInstaller(installerEntry.Installers);
				if(eventlogInstaller != null)
					{
					return eventlogInstaller;
					}
				}
			return null;
			}


		}
	}
