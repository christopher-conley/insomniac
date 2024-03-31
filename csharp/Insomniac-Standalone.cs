// Simple PowerShell host created by Ingo Karstein (http://blog.karstein-consulting.com)
// Reworked and GUI support by Markus Scholtes

using System;
using System.Collections.Generic;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Globalization;
using System.Management.Automation.Host;
using System.Security;
using System.Reflection;
using System.Runtime.InteropServices;



[assembly:AssemblyTitle("Insomniac")]
[assembly:AssemblyProduct("Insomniac")]
[assembly:AssemblyCopyright("Copyright Â© 2024 Christopher Conley")]
[assembly:AssemblyTrademark("")]
[assembly:AssemblyVersion("2024.03.31.0202")]
[assembly:AssemblyFileVersion("2024.03.31.0202")]
// not displayed in details tab of properties dialog, but embedded to file
[assembly:AssemblyDescription("Stimulants for your console")]
[assembly:AssemblyCompany("")]


namespace ModuleNameSpace
{


	internal class MainModuleRawUI : PSHostRawUserInterface
	{
		const int STD_OUTPUT_HANDLE = -11;

		//CHAR_INFO struct, which was a union in the old days
		// so we want to use LayoutKind.Explicit to mimic it as closely
		// as we can
		[StructLayout(LayoutKind.Explicit)]
		public struct CHAR_INFO
		{
			[FieldOffset(0)]
			internal char UnicodeChar;
			[FieldOffset(0)]
			internal char AsciiChar;
			[FieldOffset(2)] //2 bytes seems to work properly
			internal UInt16 Attributes;
		}

		//COORD struct
		[StructLayout(LayoutKind.Sequential)]
		public struct COORD
		{
			public short X;
			public short Y;
		}

		//SMALL_RECT struct
		[StructLayout(LayoutKind.Sequential)]
		public struct SMALL_RECT
		{
			public short Left;
			public short Top;
			public short Right;
			public short Bottom;
		}

		/* Reads character and color attribute data from a rectangular block of character cells in a console screen buffer,
			 and the function writes the data to a rectangular block at a specified location in the destination buffer. */
		[DllImport("kernel32.dll", EntryPoint = "ReadConsoleOutputW", CharSet = CharSet.Unicode, SetLastError = true)]
		internal static extern bool ReadConsoleOutput(
			IntPtr hConsoleOutput,
			/* This pointer is treated as the origin of a two-dimensional array of CHAR_INFO structures
			whose size is specified by the dwBufferSize parameter.*/
			[MarshalAs(UnmanagedType.LPArray), Out] CHAR_INFO[,] lpBuffer,
			COORD dwBufferSize,
			COORD dwBufferCoord,
			ref SMALL_RECT lpReadRegion);

		/* Writes character and color attribute data to a specified rectangular block of character cells in a console screen buffer.
			The data to be written is taken from a correspondingly sized rectangular block at a specified location in the source buffer */
		[DllImport("kernel32.dll", EntryPoint = "WriteConsoleOutputW", CharSet = CharSet.Unicode, SetLastError = true)]
		internal static extern bool WriteConsoleOutput(
			IntPtr hConsoleOutput,
			/* This pointer is treated as the origin of a two-dimensional array of CHAR_INFO structures
			whose size is specified by the dwBufferSize parameter.*/
			[MarshalAs(UnmanagedType.LPArray), In] CHAR_INFO[,] lpBuffer,
			COORD dwBufferSize,
			COORD dwBufferCoord,
			ref SMALL_RECT lpWriteRegion);

		/* Moves a block of data in a screen buffer. The effects of the move can be limited by specifying a clipping rectangle, so
			the contents of the console screen buffer outside the clipping rectangle are unchanged. */
		[DllImport("kernel32.dll", SetLastError = true)]
		static extern bool ScrollConsoleScreenBuffer(
			IntPtr hConsoleOutput,
			[In] ref SMALL_RECT lpScrollRectangle,
			[In] ref SMALL_RECT lpClipRectangle,
			COORD dwDestinationOrigin,
			[In] ref CHAR_INFO lpFill);

		[DllImport("kernel32.dll", SetLastError = true)]
			static extern IntPtr GetStdHandle(int nStdHandle);

		public override ConsoleColor BackgroundColor
		{
			get
			{
				return Console.BackgroundColor;
			}
			set
			{
				Console.BackgroundColor = value;
			}
		}

		public override System.Management.Automation.Host.Size BufferSize
		{
			get
			{
				if (Console_Info.IsOutputRedirected())
					// return default value for redirection. If no valid value is returned WriteLine will not be called
					return new System.Management.Automation.Host.Size(120, 50);
				else
					return new System.Management.Automation.Host.Size(Console.BufferWidth, Console.BufferHeight);
			}
			set
			{
				Console.BufferWidth = value.Width;
				Console.BufferHeight = value.Height;
			}
		}

		public override Coordinates CursorPosition
		{
			get
			{
				return new Coordinates(Console.CursorLeft, Console.CursorTop);
			}
			set
			{
				Console.CursorTop = value.Y;
				Console.CursorLeft = value.X;
			}
		}

		public override int CursorSize
		{
			get
			{
				return Console.CursorSize;
			}
			set
			{
				Console.CursorSize = value;
			}
		}



		public override void FlushInputBuffer()
		{
			if (!Console_Info.IsInputRedirected())
			{	while (Console.KeyAvailable)
					Console.ReadKey(true);
			}
		}

		public override ConsoleColor ForegroundColor
		{
			get
			{
				return Console.ForegroundColor;
			}
			set
			{
				Console.ForegroundColor = value;
			}
		}

		public override BufferCell[,] GetBufferContents(System.Management.Automation.Host.Rectangle rectangle)
		{
			IntPtr hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);
			CHAR_INFO[,] buffer = new CHAR_INFO[rectangle.Bottom - rectangle.Top + 1, rectangle.Right - rectangle.Left + 1];
			COORD buffer_size = new COORD() {X = (short)(rectangle.Right - rectangle.Left + 1), Y = (short)(rectangle.Bottom - rectangle.Top + 1)};
			COORD buffer_index = new COORD() {X = 0, Y = 0};
			SMALL_RECT screen_rect = new SMALL_RECT() {Left = (short)rectangle.Left, Top = (short)rectangle.Top, Right = (short)rectangle.Right, Bottom = (short)rectangle.Bottom};

			ReadConsoleOutput(hStdOut, buffer, buffer_size, buffer_index, ref screen_rect);

			System.Management.Automation.Host.BufferCell[,] ScreenBuffer = new System.Management.Automation.Host.BufferCell[rectangle.Bottom - rectangle.Top + 1, rectangle.Right - rectangle.Left + 1];
			for (int y = 0; y <= rectangle.Bottom - rectangle.Top; y++)
				for (int x = 0; x <= rectangle.Right - rectangle.Left; x++)
				{
					ScreenBuffer[y,x] = new System.Management.Automation.Host.BufferCell(buffer[y,x].AsciiChar, (System.ConsoleColor)(buffer[y,x].Attributes & 0xF), (System.ConsoleColor)((buffer[y,x].Attributes & 0xF0) / 0x10), System.Management.Automation.Host.BufferCellType.Complete);
				}

			return ScreenBuffer;
		}

		public override bool KeyAvailable
		{
			get
			{
				return Console.KeyAvailable;
			}
		}

		public override System.Management.Automation.Host.Size MaxPhysicalWindowSize
		{
			get
			{
				return new System.Management.Automation.Host.Size(Console.LargestWindowWidth, Console.LargestWindowHeight);
			}
		}

		public override System.Management.Automation.Host.Size MaxWindowSize
		{
			get
			{
				return new System.Management.Automation.Host.Size(Console.BufferWidth, Console.BufferWidth);
			}
		}

		public override KeyInfo ReadKey(ReadKeyOptions options)
		{
			ConsoleKeyInfo cki = Console.ReadKey((options & ReadKeyOptions.NoEcho)!=0);

			ControlKeyStates cks = 0;
			if ((cki.Modifiers & ConsoleModifiers.Alt) != 0)
				cks |= ControlKeyStates.LeftAltPressed | ControlKeyStates.RightAltPressed;
			if ((cki.Modifiers & ConsoleModifiers.Control) != 0)
				cks |= ControlKeyStates.LeftCtrlPressed | ControlKeyStates.RightCtrlPressed;
			if ((cki.Modifiers & ConsoleModifiers.Shift) != 0)
				cks |= ControlKeyStates.ShiftPressed;
			if (Console.CapsLock)
				cks |= ControlKeyStates.CapsLockOn;
			if (Console.NumberLock)
				cks |= ControlKeyStates.NumLockOn;

			return new KeyInfo((int)cki.Key, cki.KeyChar, cks, (options & ReadKeyOptions.IncludeKeyDown)!=0);
		}

		public override void ScrollBufferContents(System.Management.Automation.Host.Rectangle source, Coordinates destination, System.Management.Automation.Host.Rectangle clip, BufferCell fill)
		{ // no destination block clipping implemented
			// clip area out of source range?
			if ((source.Left > clip.Right) || (source.Right < clip.Left) || (source.Top > clip.Bottom) || (source.Bottom < clip.Top))
			{ // clipping out of range -> nothing to do
				return;
			}

			IntPtr hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);
			SMALL_RECT lpScrollRectangle = new SMALL_RECT() {Left = (short)source.Left, Top = (short)source.Top, Right = (short)(source.Right), Bottom = (short)(source.Bottom)};
			SMALL_RECT lpClipRectangle;
			if (clip != null)
			{ lpClipRectangle = new SMALL_RECT() {Left = (short)clip.Left, Top = (short)clip.Top, Right = (short)(clip.Right), Bottom = (short)(clip.Bottom)}; }
			else
			{ lpClipRectangle = new SMALL_RECT() {Left = (short)0, Top = (short)0, Right = (short)(Console.WindowWidth - 1), Bottom = (short)(Console.WindowHeight - 1)}; }
			COORD dwDestinationOrigin = new COORD() {X = (short)(destination.X), Y = (short)(destination.Y)};
			CHAR_INFO lpFill = new CHAR_INFO() { AsciiChar = fill.Character, Attributes = (ushort)((int)(fill.ForegroundColor) + (int)(fill.BackgroundColor)*16) };

			ScrollConsoleScreenBuffer(hStdOut, ref lpScrollRectangle, ref lpClipRectangle, dwDestinationOrigin, ref lpFill);
		}

		public override void SetBufferContents(System.Management.Automation.Host.Rectangle rectangle, BufferCell fill)
		{
			// using a trick: move the buffer out of the screen, the source area gets filled with the char fill.Character
			if (rectangle.Left >= 0)
				Console.MoveBufferArea(rectangle.Left, rectangle.Top, rectangle.Right-rectangle.Left+1, rectangle.Bottom-rectangle.Top+1, BufferSize.Width, BufferSize.Height, fill.Character, fill.ForegroundColor, fill.BackgroundColor);
			else
			{ // Clear-Host: move all content off the screen
				Console.MoveBufferArea(0, 0, BufferSize.Width, BufferSize.Height, BufferSize.Width, BufferSize.Height, fill.Character, fill.ForegroundColor, fill.BackgroundColor);
			}
		}

		public override void SetBufferContents(Coordinates origin, BufferCell[,] contents)
		{
			IntPtr hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);
			CHAR_INFO[,] buffer = new CHAR_INFO[contents.GetLength(0), contents.GetLength(1)];
			COORD buffer_size = new COORD() {X = (short)(contents.GetLength(1)), Y = (short)(contents.GetLength(0))};
			COORD buffer_index = new COORD() {X = 0, Y = 0};
			SMALL_RECT screen_rect = new SMALL_RECT() {Left = (short)origin.X, Top = (short)origin.Y, Right = (short)(origin.X + contents.GetLength(1) - 1), Bottom = (short)(origin.Y + contents.GetLength(0) - 1)};

			for (int y = 0; y < contents.GetLength(0); y++)
				for (int x = 0; x < contents.GetLength(1); x++)
				{
					buffer[y,x] = new CHAR_INFO() { AsciiChar = contents[y,x].Character, Attributes = (ushort)((int)(contents[y,x].ForegroundColor) + (int)(contents[y,x].BackgroundColor)*16) };
				}

			WriteConsoleOutput(hStdOut, buffer, buffer_size, buffer_index, ref screen_rect);
		}

		public override Coordinates WindowPosition
		{
			get
			{
				Coordinates s = new Coordinates();
				s.X = Console.WindowLeft;
				s.Y = Console.WindowTop;
				return s;
			}
			set
			{
				Console.WindowLeft = value.X;
				Console.WindowTop = value.Y;
			}
		}

		public override System.Management.Automation.Host.Size WindowSize
		{
			get
			{
				System.Management.Automation.Host.Size s = new System.Management.Automation.Host.Size();
				s.Height = Console.WindowHeight;
				s.Width = Console.WindowWidth;
				return s;
			}
			set
			{
				Console.WindowWidth = value.Width;
				Console.WindowHeight = value.Height;
			}
		}

		public override string WindowTitle
		{
			get
			{
				return Console.Title;
			}
			set
			{
				Console.Title = value;
			}
		}
	}



	// define IsInputRedirected(), IsOutputRedirected() and IsErrorRedirected() here since they were introduced first with .Net 4.5
	public class Console_Info
	{
		private enum FileType : uint
		{
			FILE_TYPE_UNKNOWN = 0x0000,
			FILE_TYPE_DISK = 0x0001,
			FILE_TYPE_CHAR = 0x0002,
			FILE_TYPE_PIPE = 0x0003,
			FILE_TYPE_REMOTE = 0x8000
		}

		private enum STDHandle : uint
		{
			STD_INPUT_HANDLE = unchecked((uint)-10),
			STD_OUTPUT_HANDLE = unchecked((uint)-11),
			STD_ERROR_HANDLE = unchecked((uint)-12)
		}

		[DllImport("Kernel32.dll")]
		static private extern UIntPtr GetStdHandle(STDHandle stdHandle);

		[DllImport("Kernel32.dll")]
		static private extern FileType GetFileType(UIntPtr hFile);

		static public bool IsInputRedirected()
		{
			UIntPtr hInput = GetStdHandle(STDHandle.STD_INPUT_HANDLE);
			FileType fileType = (FileType)GetFileType(hInput);
			if ((fileType == FileType.FILE_TYPE_CHAR) || (fileType == FileType.FILE_TYPE_UNKNOWN))
				return false;
			return true;
		}

		static public bool IsOutputRedirected()
		{
			UIntPtr hOutput = GetStdHandle(STDHandle.STD_OUTPUT_HANDLE);
			FileType fileType = (FileType)GetFileType(hOutput);
			if ((fileType == FileType.FILE_TYPE_CHAR) || (fileType == FileType.FILE_TYPE_UNKNOWN))
				return false;
			return true;
		}

		static public bool IsErrorRedirected()
		{
			UIntPtr hError = GetStdHandle(STDHandle.STD_ERROR_HANDLE);
			FileType fileType = (FileType)GetFileType(hError);
			if ((fileType == FileType.FILE_TYPE_CHAR) || (fileType == FileType.FILE_TYPE_UNKNOWN))
				return false;
			return true;
		}
	}


	internal class MainModuleUI : PSHostUserInterface
	{
		private MainModuleRawUI rawUI = null;

		public ConsoleColor ErrorForegroundColor = ConsoleColor.Red;
		public ConsoleColor ErrorBackgroundColor = ConsoleColor.Black;

		public ConsoleColor WarningForegroundColor = ConsoleColor.Yellow;
		public ConsoleColor WarningBackgroundColor = ConsoleColor.Black;

		public ConsoleColor DebugForegroundColor = ConsoleColor.Yellow;
		public ConsoleColor DebugBackgroundColor = ConsoleColor.Black;

		public ConsoleColor VerboseForegroundColor = ConsoleColor.Yellow;
		public ConsoleColor VerboseBackgroundColor = ConsoleColor.Black;

		public ConsoleColor ProgressForegroundColor = ConsoleColor.Yellow;
		public ConsoleColor ProgressBackgroundColor = ConsoleColor.DarkCyan;

		public MainModuleUI() : base()
		{
			rawUI = new MainModuleRawUI();
			rawUI.ForegroundColor = Console.ForegroundColor;
			rawUI.BackgroundColor = Console.BackgroundColor;
		}

		public override Dictionary<string, PSObject> Prompt(string caption, string message, System.Collections.ObjectModel.Collection<FieldDescription> descriptions)
		{
			if (!string.IsNullOrEmpty(caption)) WriteLine(caption);
			if (!string.IsNullOrEmpty(message)) WriteLine(message);
			Dictionary<string, PSObject> ret = new Dictionary<string, PSObject>();
			foreach (FieldDescription cd in descriptions)
			{
				Type t = null;
				if (string.IsNullOrEmpty(cd.ParameterAssemblyFullName))
					t = typeof(string);
				else
					t = Type.GetType(cd.ParameterAssemblyFullName);

				if (t.IsArray)
				{
					Type elementType = t.GetElementType();
					Type genericListType = Type.GetType("System.Collections.Generic.List"+((char)0x60).ToString()+"1");
					genericListType = genericListType.MakeGenericType(new Type[] { elementType });
					ConstructorInfo constructor = genericListType.GetConstructor(BindingFlags.CreateInstance | BindingFlags.Instance | BindingFlags.Public, null, Type.EmptyTypes, null);
					object resultList = constructor.Invoke(null);

					int index = 0;
					string data = "";
					do
					{
						try
						{
							if (!string.IsNullOrEmpty(cd.Name)) Write(string.Format("{0}[{1}]: ", cd.Name, index));
							data = ReadLine();
							if (string.IsNullOrEmpty(data))
								break;

							object o = System.Convert.ChangeType(data, elementType);
							genericListType.InvokeMember("Add", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, resultList, new object[] { o });
						}
						catch (Exception e)
						{
							throw e;
						}
						index++;
					} while (true);

					System.Array retArray = (System.Array )genericListType.InvokeMember("ToArray", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, resultList, null);
					ret.Add(cd.Name, new PSObject(retArray));
				}
				else
				{
					object o = null;
					string l = null;
					try
					{
						if (t != typeof(System.Security.SecureString))
						{
							if (t != typeof(System.Management.Automation.PSCredential))
							{
								if (!string.IsNullOrEmpty(cd.Name)) Write(cd.Name);
								if (!string.IsNullOrEmpty(cd.HelpMessage)) Write(" (Type !? for help.)");
								if ((!string.IsNullOrEmpty(cd.Name)) || (!string.IsNullOrEmpty(cd.HelpMessage))) Write(": ");
								do {
									l = ReadLine();
									if (l == "!?")
										WriteLine(cd.HelpMessage);
									else
									{
										if (string.IsNullOrEmpty(l)) o = cd.DefaultValue;
										if (o == null)
										{
											try {
												o = System.Convert.ChangeType(l, t);
											}
											catch {
												Write("Wrong format, please repeat input: ");
												l = "!?";
											}
										}
									}
								} while (l == "!?");
							}
							else
							{
								PSCredential pscred = PromptForCredential("", "", "", "");
								o = pscred;
							}
						}
						else
						{
								if (!string.IsNullOrEmpty(cd.Name)) Write(string.Format("{0}: ", cd.Name));

							SecureString pwd = null;
							pwd = ReadLineAsSecureString();
							o = pwd;
						}

						ret.Add(cd.Name, new PSObject(o));
					}
					catch (Exception e)
					{
						throw e;
					}
				}
			}

			return ret;
		}

		public override int PromptForChoice(string caption, string message, System.Collections.ObjectModel.Collection<ChoiceDescription> choices, int defaultChoice)
		{
			if (!string.IsNullOrEmpty(caption)) WriteLine(caption);
			WriteLine(message);
			do {
				int idx = 0;
				SortedList<string, int> res = new SortedList<string, int>();
				string defkey = "";
				foreach (ChoiceDescription cd in choices)
				{
					string lkey = cd.Label.Substring(0, 1), ltext = cd.Label;
					int pos = cd.Label.IndexOf('&');
					if (pos > -1)
					{
						lkey = cd.Label.Substring(pos + 1, 1).ToUpper();
						if (pos > 0)
							ltext = cd.Label.Substring(0, pos) + cd.Label.Substring(pos + 1);
						else
							ltext = cd.Label.Substring(1);
					}
					res.Add(lkey.ToLower(), idx);

					if (idx > 0) Write("  ");
					if (idx == defaultChoice)
					{
						Write(VerboseForegroundColor, rawUI.BackgroundColor, string.Format("[{0}] {1}", lkey, ltext));
						defkey = lkey;
					}
					else
						Write(rawUI.ForegroundColor, rawUI.BackgroundColor, string.Format("[{0}] {1}", lkey, ltext));
					idx++;
				}
				Write(rawUI.ForegroundColor, rawUI.BackgroundColor, string.Format("  [?] Help (default is \"{0}\"): ", defkey));

				string inpkey = "";
				try
				{
					inpkey = Console.ReadLine().ToLower();
					if (res.ContainsKey(inpkey)) return res[inpkey];
					if (string.IsNullOrEmpty(inpkey)) return defaultChoice;
				}
				catch { }
				if (inpkey == "?")
				{
					foreach (ChoiceDescription cd in choices)
					{
						string lkey = cd.Label.Substring(0, 1);
						int pos = cd.Label.IndexOf('&');
						if (pos > -1) lkey = cd.Label.Substring(pos + 1, 1).ToUpper();
						if (!string.IsNullOrEmpty(cd.HelpMessage))
							WriteLine(rawUI.ForegroundColor, rawUI.BackgroundColor, string.Format("{0} - {1}", lkey, cd.HelpMessage));
						else
							WriteLine(rawUI.ForegroundColor, rawUI.BackgroundColor, string.Format("{0} -", lkey));
					}
				}
			} while (true);
		}

		public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName, PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options)
		{
			if (!string.IsNullOrEmpty(caption)) WriteLine(caption);
			WriteLine(message);

			string un;
			if ((string.IsNullOrEmpty(userName)) || ((options & PSCredentialUIOptions.ReadOnlyUserName) == 0))
			{
				Write("User name: ");
				un = ReadLine();
			}
			else
			{
				Write("User name: ");
				if (!string.IsNullOrEmpty(targetName)) Write(targetName + "\\");
				WriteLine(userName);
				un = userName;
			}
			SecureString pwd = null;
			Write("Password: ");
			pwd = ReadLineAsSecureString();

			if (string.IsNullOrEmpty(un)) un = "<NOUSER>";
			if (!string.IsNullOrEmpty(targetName))
			{
				if (un.IndexOf('\\') < 0)
					un = targetName + "\\" + un;
			}

			PSCredential c2 = new PSCredential(un, pwd);
			return c2;
		}

		public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName)
		{
			if (!string.IsNullOrEmpty(caption)) WriteLine(caption);
			WriteLine(message);

			string un;
			if (string.IsNullOrEmpty(userName))
			{
				Write("User name: ");
				un = ReadLine();
			}
			else
			{
				Write("User name: ");
				if (!string.IsNullOrEmpty(targetName)) Write(targetName + "\\");
				WriteLine(userName);
				un = userName;
			}
			SecureString pwd = null;
			Write("Password: ");
			pwd = ReadLineAsSecureString();

			if (string.IsNullOrEmpty(un)) un = "<NOUSER>";
			if (!string.IsNullOrEmpty(targetName))
			{
				if (un.IndexOf('\\') < 0)
					un = targetName + "\\" + un;
			}

			PSCredential c2 = new PSCredential(un, pwd);
			return c2;
		}

		public override PSHostRawUserInterface RawUI
		{
			get
			{
				return rawUI;
			}
		}



		public override string ReadLine()
		{
			return Console.ReadLine();

		}

		private System.Security.SecureString getPassword()
		{
			System.Security.SecureString pwd = new System.Security.SecureString();
			while (true)
			{
				ConsoleKeyInfo i = Console.ReadKey(true);
				if (i.Key == ConsoleKey.Enter)
				{
					Console.WriteLine();
					break;
				}
				else if (i.Key == ConsoleKey.Backspace)
				{
					if (pwd.Length > 0)
					{
						pwd.RemoveAt(pwd.Length - 1);
						Console.Write("\b \b");
					}
				}
				else if (i.KeyChar != '\u0000')
				{
					pwd.AppendChar(i.KeyChar);
					Console.Write("*");
				}
			}
			return pwd;
		}

		public override System.Security.SecureString ReadLineAsSecureString()
		{
			System.Security.SecureString secstr = new System.Security.SecureString();
			secstr = getPassword();

			return secstr;
		}

		// called by Write-Host
		public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value)
		{
			ConsoleColor fgc = Console.ForegroundColor, bgc = Console.BackgroundColor;
			Console.ForegroundColor = foregroundColor;
			Console.BackgroundColor = backgroundColor;
			Console.Write(value);
			Console.ForegroundColor = fgc;
			Console.BackgroundColor = bgc;
		}

		public override void Write(string value)
		{
			Console.Write(value);
		}

		// called by Write-Debug
		public override void WriteDebugLine(string message)
		{
			WriteLineInternal(DebugForegroundColor, DebugBackgroundColor, string.Format("DEBUG: {0}", message));
		}

		// called by Write-Error
		public override void WriteErrorLine(string value)
		{
			if (Console_Info.IsErrorRedirected())
				Console.Error.WriteLine(string.Format("ERROR: {0}", value));
			else
				WriteLineInternal(ErrorForegroundColor, ErrorBackgroundColor, string.Format("ERROR: {0}", value));
		}

		public override void WriteLine()
		{
			Console.WriteLine();
		}

		public override void WriteLine(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value)
		{
			ConsoleColor fgc = Console.ForegroundColor, bgc = Console.BackgroundColor;
			Console.ForegroundColor = foregroundColor;
			Console.BackgroundColor = backgroundColor;
			Console.WriteLine(value);
			Console.ForegroundColor = fgc;
			Console.BackgroundColor = bgc;
		}

		private void WriteLineInternal(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value)
		{
			ConsoleColor fgc = Console.ForegroundColor, bgc = Console.BackgroundColor;
			Console.ForegroundColor = foregroundColor;
			Console.BackgroundColor = backgroundColor;
			Console.WriteLine(value);
			Console.ForegroundColor = fgc;
			Console.BackgroundColor = bgc;
		}

		// called by Write-Output
		public override void WriteLine(string value)
		{
			Console.WriteLine(value);
		}


		public override void WriteProgress(long sourceId, ProgressRecord record)
		{

		}

		// called by Write-Verbose
		public override void WriteVerboseLine(string message)
		{
			WriteLine(VerboseForegroundColor, VerboseBackgroundColor, string.Format("VERBOSE: {0}", message));
		}

		// called by Write-Warning
		public override void WriteWarningLine(string message)
		{
			WriteLineInternal(WarningForegroundColor, WarningBackgroundColor, string.Format("WARNING: {0}", message));
		}
	}

	internal class MainModule : PSHost
	{
		private MainAppInterface parent;
		private MainModuleUI ui = null;

		private CultureInfo originalCultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;

		private CultureInfo originalUICultureInfo = System.Threading.Thread.CurrentThread.CurrentUICulture;

		private Guid myId = Guid.NewGuid();

		public MainModule(MainAppInterface app, MainModuleUI ui)
		{
			this.parent = app;
			this.ui = ui;
		}

		public class ConsoleColorProxy
		{
			private MainModuleUI _ui;

			public ConsoleColorProxy(MainModuleUI ui)
			{
				if (ui == null) throw new ArgumentNullException("ui");
				_ui = ui;
			}

			public ConsoleColor ErrorForegroundColor
			{
				get
				{ return _ui.ErrorForegroundColor; }
				set
				{ _ui.ErrorForegroundColor = value; }
			}

			public ConsoleColor ErrorBackgroundColor
			{
				get
				{ return _ui.ErrorBackgroundColor; }
				set
				{ _ui.ErrorBackgroundColor = value; }
			}

			public ConsoleColor WarningForegroundColor
			{
				get
				{ return _ui.WarningForegroundColor; }
				set
				{ _ui.WarningForegroundColor = value; }
			}

			public ConsoleColor WarningBackgroundColor
			{
				get
				{ return _ui.WarningBackgroundColor; }
				set
				{ _ui.WarningBackgroundColor = value; }
			}

			public ConsoleColor DebugForegroundColor
			{
				get
				{ return _ui.DebugForegroundColor; }
				set
				{ _ui.DebugForegroundColor = value; }
			}

			public ConsoleColor DebugBackgroundColor
			{
				get
				{ return _ui.DebugBackgroundColor; }
				set
				{ _ui.DebugBackgroundColor = value; }
			}

			public ConsoleColor VerboseForegroundColor
			{
				get
				{ return _ui.VerboseForegroundColor; }
				set
				{ _ui.VerboseForegroundColor = value; }
			}

			public ConsoleColor VerboseBackgroundColor
			{
				get
				{ return _ui.VerboseBackgroundColor; }
				set
				{ _ui.VerboseBackgroundColor = value; }
			}

			public ConsoleColor ProgressForegroundColor
			{
				get
				{ return _ui.ProgressForegroundColor; }
				set
				{ _ui.ProgressForegroundColor = value; }
			}

			public ConsoleColor ProgressBackgroundColor
			{
				get
				{ return _ui.ProgressBackgroundColor; }
				set
				{ _ui.ProgressBackgroundColor = value; }
			}
		}

		public override PSObject PrivateData
		{
			get
			{
				if (ui == null) return null;
				return _consoleColorProxy ?? (_consoleColorProxy = PSObject.AsPSObject(new ConsoleColorProxy(ui)));
			}
		}

		private PSObject _consoleColorProxy;

		public override System.Globalization.CultureInfo CurrentCulture
		{
			get
			{
				return this.originalCultureInfo;
			}
		}

		public override System.Globalization.CultureInfo CurrentUICulture
		{
			get
			{
				return this.originalUICultureInfo;
			}
		}

		public override Guid InstanceId
		{
			get
			{
				return this.myId;
			}
		}

		public override string Name
		{
			get
			{
				return "PSRunspace-Host";
			}
		}

		public override PSHostUserInterface UI
		{
			get
			{
				return ui;
			}
		}

		public override Version Version
		{
			get
			{
				return new Version(0, 5, 0, 29);
			}
		}

		public override void EnterNestedPrompt()
		{
		}

		public override void ExitNestedPrompt()
		{
		}

		public override void NotifyBeginApplication()
		{
			return;
		}

		public override void NotifyEndApplication()
		{
			return;
		}

		public override void SetShouldExit(int exitCode)
		{
			this.parent.ShouldExit = true;
			this.parent.ExitCode = exitCode;
		}
	}

	internal interface MainAppInterface
	{
		bool ShouldExit { get; set; }
		int ExitCode { get; set; }
	}

	internal class MainApp : MainAppInterface
	{
		private bool shouldExit;

		private int exitCode;

		public bool ShouldExit
		{
			get { return this.shouldExit; }
			set { this.shouldExit = value; }
		}

		public int ExitCode
		{
			get { return this.exitCode; }
			set { this.exitCode = value; }
		}

		[MTAThread]
		private static int Main(string[] args)
		{
			System.Console.OutputEncoding = new System.Text.UnicodeEncoding();
			

			
			MainApp me = new MainApp();

			bool paramWait = false;
			string extractFN = string.Empty;

			MainModuleUI ui = new MainModuleUI();
			MainModule host = new MainModule(me, ui);
			System.Threading.ManualResetEvent mre = new System.Threading.ManualResetEvent(false);

			AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

			try
			{
				using (Runspace myRunSpace = RunspaceFactory.CreateRunspace(host))
				{
					myRunSpace.ApartmentState = System.Threading.ApartmentState.MTA;
					myRunSpace.Open();

					using (PowerShell posh = PowerShell.Create())
					{
						Console.CancelKeyPress += new ConsoleCancelEventHandler(delegate(object sender, ConsoleCancelEventArgs e)
						{
							try
							{
								posh.BeginStop(new AsyncCallback(delegate(IAsyncResult r)
								{
									mre.Set();
									e.Cancel = true;
								}), null);
							}
							catch
							{
							};
						});

						posh.Runspace = myRunSpace;
						posh.Streams.Error.DataAdded += new EventHandler<DataAddedEventArgs>(delegate(object sender, DataAddedEventArgs e)
						{
							ui.WriteErrorLine(((PSDataCollection<ErrorRecord>)sender)[e.Index].Exception.Message);
						});

						PSDataCollection<string> colInput = new PSDataCollection<string>();
						if (Console_Info.IsInputRedirected())
						{ // read standard input
							string sItem = "";
							while ((sItem = Console.ReadLine()) != null)
							{ // add to powershell pipeline
								colInput.Add(sItem);
							}
						}
						colInput.Complete();

						PSDataCollection<PSObject> colOutput = new PSDataCollection<PSObject>();
						colOutput.DataAdded += new EventHandler<DataAddedEventArgs>(delegate(object sender, DataAddedEventArgs e)
						{
							ui.WriteLine(((PSDataCollection<PSObject>)sender)[e.Index].ToString());
						});

						int separator = 0;
						int idx = 0;
						foreach (string s in args)
						{
							if (string.Compare(s, "-wait", true) == 0)
								paramWait = true;
							else if (s.StartsWith("-extract", StringComparison.InvariantCultureIgnoreCase))
							{
								string[] s1 = s.Split(new string[] { ":" }, 2, StringSplitOptions.RemoveEmptyEntries);
								if (s1.Length != 2)
								{
									Console.WriteLine("If you specify the -extract option you need to add a file for extraction in this way\r\n   -extract:\"<filename>\"");
									return 1;
								}
								extractFN = s1[1].Trim(new char[] { '\"' });
							}
							else if (string.Compare(s, "-end", true) == 0)
							{
								separator = idx + 1;
								break;
							}
							else if (string.Compare(s, "-debug", true) == 0)
							{
								System.Diagnostics.Debugger.Launch();
								break;
							}
							idx++;
						}

						Assembly executingAssembly = Assembly.GetExecutingAssembly();
						using (System.IO.Stream scriptstream = executingAssembly.GetManifestResourceStream("Insomniac-Standalone.ps1"))
						{
							using (System.IO.StreamReader scriptreader = new System.IO.StreamReader(scriptstream, System.Text.Encoding.UTF8))
							{
								string script = scriptreader.ReadToEnd();

								if (!string.IsNullOrEmpty(extractFN))
								{
									System.IO.File.WriteAllText(extractFN, script);
									return 0;
								}

								posh.AddScript(script);
							}
						}

						// parse parameters
						string argbuffer = null;
						// regex for named parameters
						System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"^-([^: ]+)[ :]?([^:]*)$");

						for (int i = separator; i < args.Length; i++)
						{
							System.Text.RegularExpressions.Match match = regex.Match(args[i]);
							double dummy;

							if ((match.Success && match.Groups.Count == 3) && (!Double.TryParse(args[i], out dummy)))
							{ // parameter in powershell style, means named parameter found
								if (argbuffer != null) // already a named parameter in buffer, then flush it
									posh.AddParameter(argbuffer);

								if (match.Groups[2].Value.Trim() == "")
								{ // store named parameter in buffer
									argbuffer = match.Groups[1].Value;
								}
								else
									// caution: when called in powershell True gets converted, when called in cmd.exe not
									if ((match.Groups[2].Value == "True") || (match.Groups[2].Value.ToUpper() == "\x24TRUE"))
									{ // switch found
										posh.AddParameter(match.Groups[1].Value, true);
										argbuffer = null;
									}
									else
										// caution: when called in powershell False gets converted, when called in cmd.exe not
										if ((match.Groups[2].Value == "False") || (match.Groups[2].Value.ToUpper() == "\x24"+"FALSE"))
										{ // switch found
											posh.AddParameter(match.Groups[1].Value, false);
											argbuffer = null;
										}
										else
										{ // named parameter with value found
											posh.AddParameter(match.Groups[1].Value, match.Groups[2].Value);
											argbuffer = null;
										}
							}
							else
							{ // unnamed parameter found
								if (argbuffer != null)
								{ // already a named parameter in buffer, so this is the value
									posh.AddParameter(argbuffer, args[i]);
									argbuffer = null;
								}
								else
								{ // position parameter found
									posh.AddArgument(args[i]);
								}
							}
						}

						if (argbuffer != null) posh.AddParameter(argbuffer); // flush parameter buffer...

						// convert output to strings
						posh.AddCommand("Out-String");
						// with a single string per line
						posh.AddParameter("Stream");

						posh.BeginInvoke<string, PSObject>(colInput, colOutput, null, new AsyncCallback(delegate(IAsyncResult ar)
						{
							if (ar.IsCompleted)
								mre.Set();
						}), null);

						while (!me.ShouldExit && !mre.WaitOne(100))
						{ };

						posh.Stop();

						if (posh.InvocationStateInfo.State == PSInvocationState.Failed)
							ui.WriteErrorLine(posh.InvocationStateInfo.Reason.Message);
					}

					myRunSpace.Close();
				}
			}
			catch (Exception ex)
			{
				Console.Write("An exception occured: ");
				Console.WriteLine(ex.Message);
			}

			if (paramWait)
			{
				Console.WriteLine("Hit any key to exit...");
				Console.ReadKey();
			}
			return me.ExitCode;
		}

		static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
		{
			throw new Exception("Unhandled exception in " + System.AppDomain.CurrentDomain.FriendlyName);
		}
	}
}