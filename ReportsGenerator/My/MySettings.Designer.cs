using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

namespace ReportsGenerator.My
{
	[GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "12.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	[CompilerGenerated]
	internal sealed partial class MySettings : ApplicationSettingsBase
	{
		[DebuggerNonUserCode]
		[EditorBrowsable(EditorBrowsableState.Advanced)]
		private static void AutoSaveSettings(object sender, EventArgs e)
		{
			if (MyProject.Application.SaveMySettingsOnExit)
			{
				MySettingsProperty.Settings.Save();
			}
		}

		public static MySettings Default
		{
			get
			{
				if (!MySettings.addedHandler)
				{
					object obj = MySettings.addedHandlerLockObject;
					ObjectFlowControl.CheckForSyncLockOnValueType(obj);
					lock (obj)
					{
						if (!MySettings.addedHandler)
						{
							MyProject.Application.Shutdown += MySettings.AutoSaveSettings;
							MySettings.addedHandler = true;
						}
					}
				}
				return MySettings.defaultInstance;
			}
		}

		[UserScopedSetting]
		[DefaultSettingValue("")]
		[DebuggerNonUserCode]
		public string Project
		{
			get
			{
				return Conversions.ToString(this["Project"]);
			}
			set
			{
				this["Project"] = value;
			}
		}

		[UserScopedSetting]
		[DebuggerNonUserCode]
		[DefaultSettingValue("")]
		public string Block
		{
			get
			{
				return Conversions.ToString(this["Block"]);
			}
			set
			{
				this["Block"] = value;
			}
		}

		[UserScopedSetting]
		[DefaultSettingValue("")]
		[DebuggerNonUserCode]
		public string Draw
		{
			get
			{
				return Conversions.ToString(this["Draw"]);
			}
			set
			{
				this["Draw"] = value;
			}
		}

		[DefaultSettingValue("E:\\123")]
		[UserScopedSetting]
		[DebuggerNonUserCode]
		public string WorkDir
		{
			get
			{
				return Conversions.ToString(this["WorkDir"]);
			}
			set
			{
				this["WorkDir"] = value;
			}
		}

		[DebuggerNonUserCode]
		[UserScopedSetting]
		[DefaultSettingValue("")]
		public string QualityList
		{
			get
			{
				return Conversions.ToString(this["QualityList"]);
			}
			set
			{
				this["QualityList"] = value;
			}
		}

		[DefaultSettingValue("12000 x 2000,8000 x 2000,6000 x 2000,6000 x 1600")]
		[UserScopedSetting]
		[DebuggerNonUserCode]
		public string NestSizeList
		{
			get
			{
				return Conversions.ToString(this["NestSizeList"]);
			}
			set
			{
				this["NestSizeList"] = value;
			}
		}

		[DefaultSettingValue("Генератор отчётов")]
		[DebuggerNonUserCode]
		[UserScopedSetting]
		public string FormTitle
		{
			get
			{
				return Conversions.ToString(this["FormTitle"]);
			}
			set
			{
				this["FormTitle"] = value;
			}
		}

		private static MySettings defaultInstance = (MySettings)SettingsBase.Synchronized(new MySettings());

		private static bool addedHandler;

		private static object addedHandlerLockObject = RuntimeHelpers.GetObjectValue(new object());
	}
}
