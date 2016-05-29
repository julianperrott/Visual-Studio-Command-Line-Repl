namespace CommandLineRepl
{
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using System;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Security.Permissions;
    using System.Text;
    using System.Threading;
    using BOOL = System.Boolean;
    using DWORD = System.UInt32;

    // unmanaged type aliases
    using HANDLE = System.IntPtr;
    using WORD = System.UInt16;

    public delegate void ReceivedOutputEventHandler(string output);

    public unsafe class ConsoleRedirect
    {
        [SecurityPermission(SecurityAction.Demand, ControlPrincipal = true)]
        private unsafe class win32api
        {
            // constants

            public const DWORD DUPLICATE_SAME_ACCESS = 0x00000002;
            public const DWORD INFINITE = 0xFFFFFFFF;
            public const DWORD STARTF_USESTDHANDLES = 0x00000100;
            public const DWORD STARTF_USESHOWWINDOW = 0x00000001;
            public const WORD SW_HIDE = 0x0000;
            public const DWORD CREATE_NEW_CONSOLE = 0x00000010;
            public const DWORD ERROR_BROKEN_PIPE = 0x0000006D;
            public const DWORD STD_INPUT_HANDLE = unchecked((DWORD)(-10));
            public const DWORD STD_OUTPUT_HANDLE = unchecked((DWORD)(-11));
            public const DWORD STD_ERROR_HANDLE = unchecked((DWORD)(-12));


            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
            public struct SECURITY_ATTRIBUTES
            {
                public DWORD nLength;
                public void* lpSecurityDescriptor;
                public BOOL bInheritHandle;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct STARTUPINFO
            {
                public DWORD cb;

                [MarshalAs(UnmanagedType.LPTStr)]
                public string lpReserved;

                [MarshalAs(UnmanagedType.LPTStr)]
                public string lpDesktop;

                [MarshalAs(UnmanagedType.LPTStr)]
                public string lpTitle;

                public DWORD dwX;
                public DWORD dwY;
                public DWORD dwXSize;
                public DWORD dwYSize;
                public DWORD dwXCountChars;
                public DWORD dwYCountChars;
                public DWORD dwFillAttribute;
                public DWORD dwFlags;
                public WORD wShowWindow;
                public WORD cbReserved2;
                public HANDLE lpReserved2;
                public HANDLE hStdInput;
                public HANDLE hStdOutput;
                public HANDLE hStdError;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
            public struct OVERLAPPED
            {
                public DWORD Internal;
                public DWORD InternalHigh;
                public DWORD Offset;
                public DWORD OffsetHigh;
                public HANDLE hEvent;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
            public struct PROCESS_INFORMATION
            {
                public HANDLE hProcess;
                public HANDLE hThread;
                public DWORD dwProcessId;
                public DWORD dwThreadId;
            }

            [DllImport("kernel32.dll", EntryPoint = "CreateProcessW", ExactSpelling = true, CharSet = CharSet.Auto)]
            public static extern BOOL CreateProcess(
                string lpApplicationName,
                string lpCommandLine,
                IntPtr lpProcessAttributes,
                IntPtr lpThreadAttributes,
                BOOL bInheritHandles,
                DWORD dwCreationFlags,
                IntPtr lpEnvironment,
                string lpCurrentDirectory,
                ref STARTUPINFO lpStartupInfo,
                out PROCESS_INFORMATION lpProcessInformation
            );

            [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
            public static extern BOOL TerminateProcess(
                HANDLE hProcess,
                WORD uExitCode
            );

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern HANDLE GetStdHandle(
                DWORD nStdHandle
            );

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern BOOL CreatePipe(
                HANDLE* hReadPipe,
                HANDLE* hWritePipe,
                ref SECURITY_ATTRIBUTES lpPipeAttributes,
                DWORD nSize
            );

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern HANDLE GetCurrentProcess();

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern BOOL WriteFile(
                HANDLE hFile,
                byte* lpBuffer,
                DWORD nNumberOfBytesToWrite,
                DWORD* lpNumberOfBytesWritten,
                OVERLAPPED* lpOverlapped
            );

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern BOOL ReadFile(
                HANDLE hFile,
                byte* lpBuffer,
                DWORD nNumberOfBytesToRead,
                DWORD* lpNumberOfBytesRead,
                OVERLAPPED* lpOverlapped
            );

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern BOOL FlushFileBuffers(
                HANDLE hFile
            );

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern BOOL CloseHandle(
                HANDLE hObject
            );

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern BOOL DuplicateHandle(
                HANDLE hSourceProcessHandle,
                HANDLE hSourceHandle,
                HANDLE hTargetProcessHandle,
                HANDLE* lpTargetHandle,
                DWORD dwDesiredAccess,
                BOOL bInheritHandle,
                DWORD dwOptions
            );
        }

        public event ReceivedOutputEventHandler ReceivedOutput;

        private string app;
        private string args;
        private Thread ReaderThread;
        private readonly IServiceProvider provider;

        public ConsoleRedirect(string app, IServiceProvider provider)
            : this(app, "", provider)
        {
           
        }

        public ConsoleRedirect(string app, string args, IServiceProvider provider)
        {
            this.app = app;
            this.args = args;
            this.provider = provider;
        }

        private HANDLE hOutputRead, hOutputWrite;
        private HANDLE hInputRead, hInputWrite;
        private HANDLE hErrorWrite;

        private HANDLE hChildProcess = HANDLE.Zero;
        private DWORD hChildProcessId = DWORD.MinValue;

        public void Start()
        {


            HANDLE hOutputReadTmp, hInputWriteTmp;
            win32api.SECURITY_ATTRIBUTES sa;
            HANDLE CurrentProcess = win32api.GetCurrentProcess();

            // Set up the security attributes struct.
            sa.nLength = (DWORD)sizeof(win32api.SECURITY_ATTRIBUTES);
            sa.lpSecurityDescriptor = null;
            sa.bInheritHandle = true;

            // Create the child output pipe.
            fixed (HANDLE* hOutputWrite_p = &hOutputWrite)
                if (!win32api.CreatePipe(&hOutputReadTmp, hOutputWrite_p, ref sa, 0))
                throw new Exception("Unable to create child output pipe");

            // Create a duplicate of the output write handle for the std error
            // write handle. This is necessary in case the child application
            // closes one of its std output handles.
            fixed (HANDLE* hErrorWrite_p = &hErrorWrite)
                if (!win32api.DuplicateHandle(CurrentProcess, hOutputWrite,
                    CurrentProcess, hErrorWrite_p,
                    0, true, win32api.DUPLICATE_SAME_ACCESS))
                throw new Exception("Unable to duplicate output handle for child's error output");

            // Create the child input pipe.
            fixed (HANDLE* hInputRead_p = &hInputRead)
                if (!win32api.CreatePipe(hInputRead_p, &hInputWriteTmp, ref sa, 0))
                throw new Exception("Unable to create child input pipe");

            // Create new output read handle and the input write handles. Set
            // the Properties to FALSE. Otherwise, the child inherits the
            // properties and, as a result, non-closeable handles to the pipes
            // are created.
            fixed (HANDLE* hOutputRead_p = &hOutputRead)
                if (!win32api.DuplicateHandle(CurrentProcess, hOutputReadTmp,
                    CurrentProcess, hOutputRead_p,
                    0, false, win32api.DUPLICATE_SAME_ACCESS))
                throw new Exception("Unable to duplicate OutputRead handle");

            fixed (HANDLE* hInputWrite_p = &hInputWrite)
                if (!win32api.DuplicateHandle(CurrentProcess, hInputWriteTmp,
                                                CurrentProcess, hInputWrite_p,
                                                0, false, win32api.DUPLICATE_SAME_ACCESS))
                throw new Exception("Unable to duplication OutputWrite handle");

            // Close inheritable copies of the handles you do not want to be
            // inherited.
            if (!win32api.CloseHandle(hOutputReadTmp))
                throw new Exception("Unable to close temporary OutputRead handle");
            if (!win32api.CloseHandle(hInputWriteTmp))
                throw new Exception("Unable to close temporary InputWrite handle");

            // start the child process
            StartRedirectedChild(hOutputWrite, hInputRead, hErrorWrite);

            // Close pipe handles (do not continue to modify the parent).
            // You need to make sure that no handles to the write end of the
            // output pipe are maintained in this process or else the pipe will
            // not close when the child process exits and the ReadFile will hang.
            if (!win32api.CloseHandle(hOutputWrite))
                throw new Exception("Unable to close the OutputWrite handle");
            if (!win32api.CloseHandle(hInputRead))
                throw new Exception("Unable to close the InputRead handle");
            if (!win32api.CloseHandle(hErrorWrite))
                throw new Exception("Unable to close the ErrorWrite handle");

            ReaderThread = new Thread(new ThreadStart(ReadOutput));
            ReaderThread.Start();
        }

        public void Stop()
        {
            win32api.TerminateProcess(hChildProcess, (WORD)0); // causes thread to exit
            //win32api.CloseHandle(hOutputRead);
            //win32api.CloseHandle(hInputWrite);
        }

        private void StartRedirectedChild(HANDLE hChildStdOut, HANDLE hChildStdIn, HANDLE hChildStdErr)
        {
            win32api.STARTUPINFO si = new win32api.STARTUPINFO();
            win32api.PROCESS_INFORMATION pi = new win32api.PROCESS_INFORMATION();

            si.cb = (DWORD)128;
            si.dwFlags = win32api.STARTF_USESTDHANDLES | win32api.STARTF_USESHOWWINDOW;
            si.hStdOutput = hChildStdOut;
            si.hStdInput = hChildStdIn;
            si.hStdError = hChildStdErr;
            si.wShowWindow = win32api.SW_HIDE; // hide the child

            // launch the process that you want to redirect
            if (!win32api.CreateProcess(null, app + " " + args, IntPtr.Zero, IntPtr.Zero,
                                        true, win32api.CREATE_NEW_CONSOLE, IntPtr.Zero, ".", ref si, out pi))
                throw new Exception("Unable to launch child process");

            // Set global child process handle to cause threads to exit.
            hChildProcess = pi.hProcess;
            hChildProcessId = pi.dwProcessId;

            // close the unncessary thread handle
            if (!win32api.CloseHandle(pi.hThread))
                throw new Exception("unable to close child thread handle");
        }

        private void ReadOutput()
        {
            const int size = 256;
            byte* buffer = stackalloc byte[size];
            DWORD dwBytesRead;
            while (true)
            {
                if (!win32api.ReadFile(hOutputRead, buffer, (DWORD)size, &dwBytesRead, null) || dwBytesRead == 0)
                    break; // pipe done -- normal exit path
                if (ReceivedOutput != null)
                {
                    ReceivedOutput(_getstring(buffer, dwBytesRead));
                }
            }
        }

        private string _getstring(byte* buffer, long count)
        {
            StringBuilder sb = new StringBuilder();
            for (; count > 0; count--) sb.Append((char)*buffer++);
            return sb.ToString();
        }

        public void SendInput(string input)
        {
            byte[] buffer = ASCIIEncoding.ASCII.GetBytes(input);
            if (hChildProcess != HANDLE.Zero)
            {
                if (input.StartsWith("^c")|| input.StartsWith("^C") || input.Equals("kill"))
                {
                    Process[] runningProcesses = Process.GetProcesses();
                    foreach (Process process in runningProcesses)
                    {
                        if(process.Id == hChildProcessId)
                        {
                            process.Kill();
                        }
                    }

                    Stop();
                    Start();
                }

                else
                {
                    DWORD dwBytesWritten;
                    fixed (byte* input_p = &buffer[0])
                    {
                        win32api.WriteFile(hInputWrite, input_p, (DWORD)input.Length, &dwBytesWritten, null);
                        win32api.FlushFileBuffers(hInputWrite);
                    }
                }
            }
        }
        private void ShowMessageBox(string message)
        {
            VsShellUtilities.ShowMessageBox(this.provider, message, "", OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}