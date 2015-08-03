// TODO: mouse over text in system tray
// TODO: Determine what the current time zone is
//       HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation
// TODO: Create "city manager" to allow assignment of cities to time zones (save XML)
// TODO: Create configurable list with zone and time info
// TODO: NTP time sync
// TODO: auto update of city list and application?
// TODO: image for DST on /off, other icons
// TODO: incorporate the israel DST rules
namespace Junkosoft
{
	#region Namespaces
	using System;
	using System.Drawing;
	using System.Collections;
	using System.ComponentModel;
	using System.Configuration;
	using System.Windows.Forms;
	using System.Data;
	using System.Text;
	using System.Runtime.InteropServices;
	using System.Diagnostics;
	using System.Xml;
	using Microsoft.Win32;
	#endregion 

	#region Public Structs
	[StructLayout(LayoutKind.Sequential,Pack=2)]
	public struct SYSTEMTIME
	{
		public short wYear;
		public short wMonth;
		public short wDayOfWeek;
		public short wDay;
		public short wHour;
		public short wMinute;
		public short wSecond;
		public short wMilliseconds;
	}

	// MS KB Q115231
	[StructLayout(LayoutKind.Sequential,Pack=2)]
	public struct TZI
	{
		public int Bias;
		public int StandardBias;
		public int DaylightBias;
		public SYSTEMTIME StandardDate;
		public SYSTEMTIME DaylightDate;
	}
	#endregion 

	/// <summary>
	/// World Clock Application
	/// </summary>
	public class WorldTime : System.Windows.Forms.Form
	{
		#region Controls
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.ComboBox TimeZoneList;
		private System.Windows.Forms.Label LocalTime;
		private System.Windows.Forms.Timer UpdateTimer;
		private System.Windows.Forms.NotifyIcon TrayIcon;
		#endregion

		#region Private Fields
		private DataTable _timeZoneData;
		private bool exitFromSystemTray;
		#endregion 

		#region Constructors
		public WorldTime()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			exitFromSystemTray = false;

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		#endregion 

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(WorldTime));
			this.TimeZoneList = new System.Windows.Forms.ComboBox();
			this.UpdateTimer = new System.Windows.Forms.Timer(this.components);
			this.LocalTime = new System.Windows.Forms.Label();
			this.TrayIcon = new System.Windows.Forms.NotifyIcon(this.components);
			this.SuspendLayout();
			// 
			// TimeZoneList
			// 
			this.TimeZoneList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.TimeZoneList.Location = new System.Drawing.Point(8, 8);
			this.TimeZoneList.Name = "TimeZoneList";
			this.TimeZoneList.Size = new System.Drawing.Size(368, 21);
			this.TimeZoneList.TabIndex = 0;
			this.TimeZoneList.SelectedIndexChanged += new System.EventHandler(this.TimeZoneList_SelectedIndexChanged);
			// 
			// UpdateTimer
			// 
			this.UpdateTimer.Enabled = true;
			this.UpdateTimer.Tick += new System.EventHandler(this.UpdateTimer_Tick);
			// 
			// LocalTime
			// 
			this.LocalTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.LocalTime.Location = new System.Drawing.Point(8, 40);
			this.LocalTime.Name = "LocalTime";
			this.LocalTime.Size = new System.Drawing.Size(368, 48);
			this.LocalTime.TabIndex = 18;
			this.LocalTime.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// TrayIcon
			// 
			this.TrayIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("TrayIcon.Icon")));
			this.TrayIcon.Text = "";
			this.TrayIcon.Visible = true;
			this.TrayIcon.MouseDown += new System.Windows.Forms.MouseEventHandler(this.TrayIcon_MouseDown);
			// 
			// WorldTime
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(384, 94);
			this.Controls.Add(this.LocalTime);
			this.Controls.Add(this.TimeZoneList);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "WorldTime";
			this.ShowInTaskbar = false;
			this.Text = "WorldTime";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.WorldTime_Closing);
			this.Load += new System.EventHandler(this.TimeZone_Load);
			this.ResumeLayout(false);

		}
		#endregion

		#region Form Events
		private void TimeZone_Load(object sender, System.EventArgs e)
		{
			// Add a menu to the system tray icon
			MenuItem[] items = new MenuItem[1];
			items[0] = new MenuItem("E&xit");
 			this.TrayIcon.ContextMenu = new ContextMenu(items);
			this.TrayIcon.ContextMenu.MenuItems[0].Click += new EventHandler(Exit_Click);

			// Create a table and add columns for all of the needed information about a time zone
			_timeZoneData = new DataTable();
			_timeZoneData.Columns.Add(new DataColumn("Display", System.Type.GetType("System.String")));
			_timeZoneData.Columns.Add(new DataColumn("Dlt", System.Type.GetType("System.String")));
			_timeZoneData.Columns.Add(new DataColumn("TZI", System.Type.GetType("Junkosoft.TZI")));
			_timeZoneData.Columns.Add(new DataColumn("Std", System.Type.GetType("System.String")));
			_timeZoneData.Columns.Add(new DataColumn("DaylightBias", System.Type.GetType("System.Int32")));
			_timeZoneData.Columns.Add(new DataColumn("StandardBias", System.Type.GetType("System.Int32")));
			_timeZoneData.Columns.Add(new DataColumn("HasDST", System.Type.GetType("System.Boolean")));
			
			// Set the primary key column for sorting
            _timeZoneData.PrimaryKey = new DataColumn[] { _timeZoneData.Columns["Display"] };

			// Load the registry keys for Windows NT, 2000, and XP
			RegistryKey tzRootKey;
			tzRootKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Time Zones");
			
			// We are running under Windows 98
			if (tzRootKey == null)
				tzRootKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Time Zones");
            
			// If the object is still null throw an exception
			if (tzRootKey == null)
			{
				throw new Exception("Unable to load time zone data from the registry");
			}
			else
			{
				// Read all of the time zone nodes and get info from each
				foreach (string currentName in tzRootKey.GetSubKeyNames())
				{
					RegistryKey currentTZKey = tzRootKey.OpenSubKey(currentName);
					DataRow newTZ = _timeZoneData.NewRow();
					newTZ["Display"] = currentTZKey.GetValue("Display");
					newTZ["Dlt"] = currentTZKey.GetValue("Dlt");
					newTZ["Std"] = currentTZKey.GetValue("Std");

					// Get the TZI data and decode it
					TZI TZIData = GetTZI((byte[])currentTZKey.GetValue("TZI"));
					newTZ["TZI"] = TZIData;
					newTZ["DaylightBias"] = TZIData.DaylightBias + TZIData.Bias;
					newTZ["StandardBias"] = TZIData.StandardBias + TZIData.Bias;

					// Determine if the time zone uses DST
					bool hasDST = (TZIData.StandardDate.wMonth > 0 && TZIData.DaylightDate.wMonth > 0);
					newTZ["HasDST"] = hasDST;
			
					// Add the row to the table and close the key
					_timeZoneData.Rows.Add(newTZ);
					currentTZKey.Close();
				}

				// Close the root key
				tzRootKey.Close();

				// Create a view to sort the data
				DataView timeZoneView = new DataView(_timeZoneData);
				timeZoneView.Sort = "StandardBias DESC";
				
				// Bind the data to the grid
				TimeZoneList.DataSource = timeZoneView;
				TimeZoneList.DisplayMember = "Display";
				TimeZoneList.ValueMember = "Display";

				// Get the previous timezone from the registry
				try
				{
					RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Junkosoft\\WorldTime", false);
					TimeZoneList.SelectedValue = (string)key.GetValue("SelectedTimeZone");
				}
				catch
				{
                    RegistryKey newKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Junkosoft\\WorldTime");
					SaveCurrentSelection();
				}
			}
			
			// Update the clock display
			UpdateTime();
		}

		private void WorldTime_Closing(object sender, CancelEventArgs e)
		{
			if (exitFromSystemTray)
			{
				SaveCurrentSelection();
			}			
			else
			{
				this.Hide();
				e.Cancel = true;
			}
		}
		#endregion

		#region Control Events
		private void TimeZoneList_SelectedIndexChanged(object sender, EventArgs e)
		{
			UpdateTime();
		}

		private void UpdateTimer_Tick(object sender, System.EventArgs e)
		{
			UpdateTime();
		}

		private void TrayIcon_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Left)
			{
				this.Show();
				this.Activate();
			}
		}

		private void Exit_Click(object sender, EventArgs e)
		{
			exitFromSystemTray = true;
			this.Close();
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		public static void Main() 
		{
			Application.Run(new WorldTime());
		}
		#endregion 

		#region Protected Methods
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		/// <summary>
		/// This method fires when when Windows system events occur (shutdown, restart, logoff, etc)
		/// </summary>
		/// <param name="m">Referenced Message object</param>
		protected override void WndProc(ref Message m)
		{
			const int WM_QUERYENDSESSION = 17;
			if (m.Msg == WM_QUERYENDSESSION)
				exitFromSystemTray = true;
    
			// If this is WM_QUERYENDSESSION, the closing event should be fired in the base WndProc
			base.WndProc(ref m);
		}
		#endregion 

		#region Private Methods
		/// <summary>
		/// Decodes the Time Zone Information structure from the byte array
		/// </summary>
		/// <param name="bytes"></param>
		/// <returns></returns>
		private TZI GetTZI(byte[] bytes)
		{
			if (bytes == null || bytes.Length < Marshal.SizeOf(typeof(TZI)))
				throw new Exception("TZI byte array is not a valid size");

			TZI currentTZI = new TZI();
			GCHandle handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
			IntPtr buffer = handle.AddrOfPinnedObject();
			currentTZI = (TZI) Marshal.PtrToStructure (buffer, typeof(TZI));
			handle.Free();
			return currentTZI;
		}

		/// <summary>
		/// Converts the SYSTEMTIME structure into a DateTime object
		/// </summary>
		/// <param name="st"></param>
		/// <param name="year"></param>
		/// <returns></returns>
		private DateTime ConvertSystemTime(SYSTEMTIME st, int year)
		{
			DateTime returnDate = FindNthDayOfMonth(year, st.wMonth, st.wDay, (DayOfWeek)st.wDayOfWeek);
			returnDate = returnDate.AddHours(st.wHour);
			returnDate = returnDate.AddMinutes(st.wMinute);
			returnDate = returnDate.AddSeconds(st.wSecond);
			returnDate = returnDate.AddMilliseconds(st.wMilliseconds);
			return returnDate;
		}

		/// <summary>
		/// Get the Nth occurance of the specified day of the week
		/// </summary>
		/// <param name="year">The specified year</param>
		/// <param name="month">The specified month</param>
		/// <param name="occurance">The Nth occurance</param>
		/// <param name="targetdayofweek">The specified day of the week</param>
		/// <returns>DateTime object containing the date</returns>
		private DateTime FindNthDayOfMonth(int year, int month, int occurance, DayOfWeek targetdayofweek)
		{
			// Find the first day of the month
			DateTime dt = new DateTime(year, month, 1);

			// Subtract day of week from Sunday to get offset to the
			// first Sunday in the month.
			// Modulo 7 is used for Sunday the 1st.
			int sun = (int)DayOfWeek.Sunday;
			int today = (int)dt.DayOfWeek;
			int ofs = (7 + sun - today) % 7;
			int day = (1 + ofs + (int)targetdayofweek) + (7 * (occurance-1));
			
			// Adjust invalid day values
			while (day < 0)
			{
				day += 7;
			}

			while (day > DateTime.DaysInMonth(year, month))
			{
				day -= 7;
			}

			dt = new DateTime(year, month, day);
			return dt;				
		}
		
		/// <summary>
		/// Determines if the is Time Zone currently on Daylight Saving Time
		/// </summary>
		/// <param name="TZIData"></param>
		/// <returns></returns>
		private bool IsOnDST(TZI TZIData)
		{
			// Get the current time (GMT)
			DateTime currentDate = DateTime.UtcNow;

			// Get the end date for the current year and convert to GMT (important!)
			DateTime standardDate = ConvertSystemTime(TZIData.StandardDate, currentDate.Year);
			DateTime standardDateGMT = standardDate.AddMinutes(TZIData.DaylightBias + TZIData.Bias);

			// Determine the current (or next) end date
			if (currentDate > standardDateGMT)
			{
				standardDate = ConvertSystemTime(TZIData.StandardDate, currentDate.Year+1);
			}
						
			// Determine the current (or next) start date
			DateTime daylightDate;
			if (TZIData.StandardDate.wMonth > TZIData.DaylightDate.wMonth)
			{
				// Northern hemisphere
				daylightDate = ConvertSystemTime(TZIData.DaylightDate, standardDate.Year);
			}
			else
			{
				// Southern hemisphere
				daylightDate = ConvertSystemTime(TZIData.DaylightDate, standardDate.Year-1);
			}

			DateTime daylightDateGMT = daylightDate.AddMinutes(TZIData.StandardBias + TZIData.Bias);
			return daylightDateGMT <= currentDate && currentDate <= standardDateGMT;
		}

		/// <summary>
		/// Updates the value in the local time label
		/// </summary>
		private void UpdateTime()
		{
            if (TimeZoneList.SelectedValue.GetType() == Type.GetType("System.String"))
			{
				// Get the currently selected row and update it if time zone uses DST
                string display = (string)TimeZoneList.SelectedValue;
                DataRow selectedRow = _timeZoneData.Rows.Find(display);

				bool onDST = false;
				if ((bool)selectedRow["HasDST"])
				{
					onDST = IsOnDST((TZI)selectedRow["TZI"]);
				}
				
				// Display the current GMT offset factoring in DST
				DateTime currentTime = DateTime.UtcNow;

				if (onDST)
				{
					LocalTime.Text = currentTime.AddMinutes(-(int)selectedRow["DaylightBias"]).ToString();
				}
				else
				{
					LocalTime.Text = currentTime.AddMinutes(-(int)selectedRow["StandardBias"]).ToString();
				}
			}
		}

		/// <summary>
		/// Creates an XML document with all of the time zone information
		/// </summary>
		/// <param name="docName"></param>
		private void CreateXmlDocument(string docName)
		{
			// Create an empty XML document and the root node
			XmlDocument doc = new System.Xml.XmlDocument();
			XmlNode parentnode = doc.CreateElement("TimeZones");
			doc.AppendChild(parentnode);

			// Get attributes from all of the data rows
			foreach (DataRow row in _timeZoneData.Rows)
			{
				// Create new element for the Time Zone and add the attributes
				XmlNode childnode = doc.CreateElement("TimeZone");
				XmlAttribute displaynode = doc.CreateAttribute("Display");
				displaynode.Value = row["Display"].ToString();
				childnode.Attributes.Append(displaynode);
				parentnode.AppendChild(childnode);

				// Create a node for the cities
				XmlNode citiesNode = doc.CreateElement("Cities");
				childnode.AppendChild(citiesNode);
			}

			// Save the XML document
			doc.Save(docName);
		}

		/// <summary>
		/// Saves the currently selected time zone
		/// </summary>
		private void SaveCurrentSelection()
		{
			RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\\Junkosoft\\WorldTime", true);
			key.SetValue("SelectedTimeZone", TimeZoneList.SelectedValue);
		}
		#endregion
	}
}
