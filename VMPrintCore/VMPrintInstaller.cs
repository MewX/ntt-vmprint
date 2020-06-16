/****************************************************************************/
/*                                                                          */
/*  Copyright (C) 2019                                                      */
/*          （株）ＮＴＴデータ                                              */
/*           第三金融事業本部 戦略ビジネス本部　システム企画担当            */
/*                                                                          */
/*  収容物  ＣＯＮＴＩＭＩＸＥ    VirtualPrinterDriver                      */
/*                                                                          */
/****************************************************************************/
/*--------------------------------------------------------------------------*/
/*〈対象業務名〉                                                            */
/*〈対象業務ＩＤ〉                                                          */
/*〈モジュール名〉                  VMPrintInstaller                        */
/*〈モジュールＩＤ〉                                                        */
/*〈モジュール通番〉                                                        */
/*--------------------------------------------------------------------------*/
/* ＜適応ＯＳ＞                     Windows 10 XXX                          */
/* ＜開発環境＞                     Microsoft Visual Studio 2017            */
/*--------------------------------------------------------------------------*/
/* ＜開発システム名＞               ＣＯＮＴＩＭＩＸＥ                      */
/* ＜開発システム番号＞                                                     */
/*--------------------------------------------------------------------------*/
/* ＜開発担当名＞                   システム企画担当                        */
/* ＜電話番号＞                     050-5546-2418                           */
/*--------------------------------------------------------------------------*/
/* ＜設計者名＞                     範                                      */
/* ＜設計年月日＞                   2019年02月19日　　　　　                */
/* ＜設計修正者名＞                                                         */
/* ＜設計修正年月日及び修正ＩＤ＞                                           */
/*--------------------------------------------------------------------------*/
/* ＜ソース作成者名＞               範                                      */
/* ＜ソース作成年月日＞             2019年02月19日　　　　　                */
/* ＜ソース修正者名＞                                                       */
/* ＜ソース修正年月日及び修正ＩＤ＞                                         */
/*--------------------------------------------------------------------------*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.Security;

using Microsoft.Win32;

namespace VMPrintCore
{
 
    public class VMPrintInstaller
    {
        #region Printer Driver Win32 API Constants

        const uint DRIVER_KERNELMODE = 0x00000001;
        const uint DRIVER_USERMODE =  0x00000002;
        const uint APD_STRICT_UPGRADE =  0x00000001;
        const uint APD_STRICT_DOWNGRADE = 0x00000002;
        const uint APD_COPY_ALL_FILES = 0x00000004;
        const uint APD_COPY_NEW_FILES = 0x00000008;
        const uint APD_COPY_FROM_DIRECTORY = 0x00000010;
        const uint DPD_DELETE_UNUSED_FILES = 0x00000001;
        const uint DPD_DELETE_SPECIFIC_VERSION = 0x00000002;
        const uint DPD_DELETE_ALL_FILES = 0x00000004;
        const uint PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST = 0x200;
        const int WIN32_FILE_ALREADY_EXISTS = 183; // Returned by XcvData "AddPort" if the port already exists
        #endregion

        private readonly TraceSource logEventSource;
        private readonly String logEventSourceNameDefault = "VMPrintCore";

        const string ENVIRONMENT_64 = "Windows x64";
        const string MONITORDLL = "redmon64vmprint.dll";
        const string PORTMONITOR = "VMPRINT";
        const string PORTNAME = "VMPrintPort:";
        const string DRIVERFILE = "PSCRIPT5.DLL";
        const string DRIVERUIFILE = "PS5UI.DLL";
        const string DRIVERHELPFILE = "PSCRIPT.HLP";
        const string DRIVERDATAFILE = "SCPDFPRN.PPD";
        
        enum DriverFileIndex
        {
            Min = 0,
            DriverFile = Min,
            UIFile,
            HelpFile,
            DataFile,
            Max = DataFile
        };

        readonly String[] printerDriverFiles = new String[] { DRIVERFILE, DRIVERUIFILE, DRIVERHELPFILE, DRIVERDATAFILE };
        readonly String[] printerDriverDependentFiles = new String[] { "PSCRIPT.NTF" };

        #region Error messages for Trace/Debug

        #endregion

        public void AddTraceListener(TraceListener additionalListener)
        {
            this.logEventSource.Listeners.Add(additionalListener);
        }

        
        #region Constructors

        public VMPrintInstaller()
        {
            this.logEventSource = new TraceSource(logEventSourceNameDefault);
            this.logEventSource.Switch = new SourceSwitch("VMPrintCoreAll");
            this.logEventSource.Switch.Level = SourceLevels.All;
        }

        #endregion

        #region Port operations

#if DEBUG
        public bool AddVMPrintPort_Test()
        {
            return AddVMPrintPort();
        }
#endif

        private bool AddVMPrintPort()
        {
            bool portAdded = false;

            int portAddResult = DoXcvDataPortOperation(PORTNAME, PORTMONITOR, "AddPort");
            switch (portAddResult)
            {
                case 0:
                case WIN32_FILE_ALREADY_EXISTS: // Port already exists - this is OK, we'll just keep using it
                    portAdded = true;
                    break;
            }
            return portAdded;
        }

        public bool DeleteVMPrintPort()
        {
            bool portDeleted = false;

            int portDeleteResult = DoXcvDataPortOperation(PORTNAME, PORTMONITOR, "DeletePort");
            switch (portDeleteResult)
            {
                case 0:
                    portDeleted = true;
                    break;
            }
            return portDeleted;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="portName"></param>
        /// <param name="xcvDataOperation"></param>
        /// <returns></returns>
        /// <remarks>I can't remember the name/link of the developer who wrote this code originally,
        /// so I can't provide a link or credit.</remarks>
        private int DoXcvDataPortOperation(string portName, string portMonitor, string xcvDataOperation)
        {

            int win32ErrorCode;

            PRINTER_DEFAULTS def = new PRINTER_DEFAULTS();

            def.pDatatype = null;
            def.pDevMode = IntPtr.Zero;
            def.DesiredAccess = 1; //Server Access Administer

            IntPtr hPrinter = IntPtr.Zero;

            //仮想プリンターが開ける場合、ポートをインストールするまたはアンインストールする
            if (NativeMethods.OpenPrinter(",XcvMonitor " + portMonitor, ref hPrinter, def) != 0)
            {
                if (!portName.EndsWith("\0"))
                    portName += "\0"; // Must be a null terminated string

                // Must get the size in bytes. Rememeber .NET strings are formed by 2-byte characters
                uint size = (uint)(portName.Length * 2);

                // Alloc memory in HGlobal to set the portName
                IntPtr portPtr = Marshal.AllocHGlobal((int)size);
                Marshal.Copy(portName.ToCharArray(), 0, portPtr, portName.Length);

                uint needed; // Not that needed in fact...
                uint xcvResult; // Will receive de result here

                NativeMethods.XcvData(hPrinter, xcvDataOperation, portPtr, size, IntPtr.Zero, 0, out needed, out xcvResult);

                NativeMethods.ClosePrinter(hPrinter);
                Marshal.FreeHGlobal(portPtr);
                win32ErrorCode = (int)xcvResult;
            }
            else
            {
                win32ErrorCode = Marshal.GetLastWin32Error();
            }
            return win32ErrorCode;

        }

        #endregion

        #region Port Monitor

        /// <summary>
        /// Adds the VM Print port monitor
        /// </summary>
        /// <param name="monitorFilePath">Directory where the uninstalled monitor dll is located</param>
        /// <returns>true if the monitor is installed, false if install failed</returns>
        public bool AddVMPrintPortMonitor(String monitorFilePath)
        {
            bool monitorAdded = false;

            IntPtr oldRedirectValue = IntPtr.Zero;

            try
            {
                oldRedirectValue = DisableWow64Redirection();

                //仮想プリンターのモニタポートが存在しない場合、インストールする、それ以外の場合、エラーとする
                if (!DoesMonitorExist(PORTMONITOR))
                {
                    // Copy the monitor DLL to
                    // the system directory
                    String fileSourcePath = Path.Combine(monitorFilePath, MONITORDLL);
                    String fileDestinationPath = Path.Combine(Environment.SystemDirectory, MONITORDLL);
                    try
                    {
                        File.Copy(fileSourcePath, fileDestinationPath, true);
                    }
                    catch (IOException)
                    {
                        // File in use, log -
                        // this is OK because it means the file is already there
                    }
                    MONITOR_INFO_2 newMonitor = new MONITOR_INFO_2();
                    newMonitor.pName = PORTMONITOR;
                    newMonitor.pEnvironment = ENVIRONMENT_64;
                    newMonitor.pDLLName = MONITORDLL;
                    if (!AddPortMonitor(newMonitor))
                        logEventSource.TraceEvent(TraceEventType.Error,
                                                  (int)TraceEventType.Error,
                                                  String.Format(Properties.Resources.PORT_NOT_ADDED, PORTMONITOR) + Environment.NewLine +
                                                  String.Format(Properties.Resources.WIN32ERROR, Marshal.GetLastWin32Error().ToString()));
                    else
                        monitorAdded = true;
                }
                else
                {
                    // Monitor already installed -
                    // log it, and keep going
                    logEventSource.TraceEvent(TraceEventType.Warning,
                                              (int)TraceEventType.Warning,
                                              String.Format(Properties.Resources.PORT_ALREADY_INSTALLED, PORTMONITOR));
                    monitorAdded = true;
                }

            }
            finally
            {
                if (oldRedirectValue != IntPtr.Zero) RevertWow64Redirection(oldRedirectValue);
            }


            return monitorAdded;
        }


        /// <summary>
        /// Disables WOW64 system directory file redirection
        /// if the current process is both
        /// 32-bit, and running on a 64-bit OS -
        /// Compiling for 64-bit OS, and setting the install dir to "ProgramFiles64"
        /// should ensure this code never runs in production
        /// </summary>
        /// <returns>A Handle, which should be retained to reenable redirection</returns>
        private IntPtr DisableWow64Redirection()
        {
            IntPtr oldValue = IntPtr.Zero;
            //64Bit OSで32Bitプログラムを動く場合、Wow64を無効する
            if (Environment.Is64BitOperatingSystem && !Environment.Is64BitProcess)
                if (!NativeMethods.Wow64DisableWow64FsRedirection(ref oldValue))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), Properties.Resources.WOW64_NOT_DISABLE);
            return oldValue;
        }

        /// <summary>
        /// Reenables WOW64 system directory file redirection
        /// if the current process is both
        /// 32-bit, and running on a 64-bit OS -
        /// Compiling for 64-bit OS, and setting the install dir to "ProgramFiles64"
        /// should ensure this code never runs in production
        /// </summary>
        /// <param name="oldValue">A Handle value - should be retained from call to <see cref="DisableWow64Redirection"/></param>
        private void RevertWow64Redirection(IntPtr oldValue)
        {
            //64Bit OSで32Bitプログラムを動く場合、Wow64を有効する
            if (Environment.Is64BitOperatingSystem && !Environment.Is64BitProcess)
            {
                if (!NativeMethods.Wow64RevertWow64FsRedirection(oldValue))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error(), Properties.Resources.WOW64_NOT_REENABLE);
                }
            }
        }

        /// <summary>
        /// Removes the VM Print port monitor
        /// </summary>
        /// <returns>true if monitor successfully removed, false if removal failed</returns>
        public bool RemoveVMPrintPortMonitor()
        {
            bool monitorRemoved = false;
            //ポートモニタを削除する
            if ((NativeMethods.DeleteMonitor(null, ENVIRONMENT_64, PORTMONITOR)) != 0)
            {
                monitorRemoved = true;
                // Try to remove the monitor DLL now
                if (!DeleteVMPrintPortMonitorDll())
                {
                    logEventSource.TraceEvent(TraceEventType.Warning,
                                              (int)TraceEventType.Warning,
                                              Properties.Resources.MONITOR_NOT_REMOVE);
                }
            }
            return monitorRemoved;
        }

        private bool DeleteVMPrintPortMonitorDll()
        {
            return DeletePortMonitorDll(MONITORDLL);
        }

        private bool DeletePortMonitorDll(String monitorDll)
        {
            bool monitorDllRemoved = false;

            String monitorDllFullPathname = String.Empty;
            IntPtr oldRedirectValue = IntPtr.Zero;
            try
            {
                oldRedirectValue = DisableWow64Redirection();

                monitorDllFullPathname = Path.Combine(Environment.SystemDirectory, monitorDll);
                
                File.Delete(monitorDllFullPathname);
                monitorDllRemoved = true;
            }
            catch (Win32Exception windows32Ex)
            {
                // This one is likely very bad -
                // log and rethrow so we don't continue
                // to try to uninstall
                logEventSource.TraceEvent(TraceEventType.Critical, 
                                          (int)TraceEventType.Critical,
                                          Properties.Resources.NATIVE_COULDNOTENABLE64REDIRECTION + String.Format(Properties.Resources.WIN32ERROR, windows32Ex.NativeErrorCode.ToString()));
                throw;
            }
            catch (IOException)
            {
                // File still in use
                logEventSource.TraceEvent(TraceEventType.Error, (int)TraceEventType.Error, String.Format(Properties.Resources.FILENOTDELETED_INUSE, monitorDllFullPathname));  
            }
            catch (UnauthorizedAccessException)
            {
                // File is readonly, or file permissions do not allow delete
                logEventSource.TraceEvent(TraceEventType.Error, (int)TraceEventType.Error, String.Format(Properties.Resources.FILENOTDELETED_INUSE, monitorDllFullPathname));
            }
            finally
            {
                try
                {
                    if (oldRedirectValue != IntPtr.Zero) RevertWow64Redirection(oldRedirectValue);
                }
                catch (Win32Exception windows32Ex)
                {
                    // Couldn't turn file redirection back on -
                    // this is not good
                    logEventSource.TraceEvent(TraceEventType.Critical, 
                                              (int)TraceEventType.Critical,
                                              Properties.Resources.NATIVE_COULDNOTREVERT64REDIRECTION + String.Format(Properties.Resources.WIN32ERROR, windows32Ex.NativeErrorCode.ToString()));
                    throw;
                }
            }

            return monitorDllRemoved;

        }

        private bool AddPortMonitor(MONITOR_INFO_2 newMonitor)
        {
            bool monitorAdded = false;
            //モニタポートを追加する
            if ((NativeMethods.AddMonitor(null, 2, ref newMonitor) != 0))
            {
                monitorAdded = true;
            }
            return monitorAdded;
        }

        private bool DeletePortMonitor(String monitorName)
        {
            bool monitorDeleted = false;
            //モニタポートを削除する
            if ((NativeMethods.DeleteMonitor(null, ENVIRONMENT_64, monitorName)) != 0)
            {
                monitorDeleted = true;
            }
            return monitorDeleted;
        }

        private bool DoesMonitorExist(String monitorName)
        {
            bool monitorExists = false;
            List<MONITOR_INFO_2> portMonitors = EnumerateMonitors();
            foreach (MONITOR_INFO_2 portMonitor in portMonitors)
            {
                //ポートの名前が一致する場合、Trueと戻す
                if (portMonitor.pName == monitorName)
                {
                    monitorExists = true;
                    break;
                }
            }
            return monitorExists;
        }


        public List<MONITOR_INFO_2> EnumerateMonitors()
        {
            List<MONITOR_INFO_2> portMonitors = new List<MONITOR_INFO_2>();

            uint pcbNeeded = 0;
            uint pcReturned = 0;

            //全てモニタの名前を取得する
            if (!NativeMethods.EnumMonitors(null, 2, IntPtr.Zero, 0, ref pcbNeeded, ref pcReturned))
            {
                IntPtr pMonitors = Marshal.AllocHGlobal((int)pcbNeeded);
                if (NativeMethods.EnumMonitors(null, 2, pMonitors, pcbNeeded, ref pcbNeeded, ref pcReturned))
                {
                    IntPtr currentMonitor = pMonitors;
                    for (int i = 0; i < pcReturned; i++)
                    {
                        portMonitors.Add((MONITOR_INFO_2)Marshal.PtrToStructure(currentMonitor, typeof(MONITOR_INFO_2)));
                        currentMonitor = IntPtr.Add(currentMonitor, Marshal.SizeOf(typeof(MONITOR_INFO_2)));
                    }
                    Marshal.FreeHGlobal(pMonitors);

                }
                else
                {
                    // Failed to retrieve enumerated monitors
                    throw new Win32Exception(Marshal.GetLastWin32Error(), Properties.Resources.MONITOR_NOT_ENUMERATE);
                }

            }
            else
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), Properties.Resources.ENUMMONITOR_ZERO_SIZE_BUFFER);
            }

            return portMonitors;
        }

        #endregion

        #region Printer Install

        public String RetrievePrinterDriverDirectory()
        {
            StringBuilder driverDirectory = new StringBuilder(1024);
            int dirSizeInBytes = 0;
            //プリンターのフォルダーを取得する
            if (!NativeMethods.GetPrinterDriverDirectory(null,
                                                         null,
                                                         1,
                                                         driverDirectory,
                                                         1024,
                                                         ref dirSizeInBytes))
                throw new DirectoryNotFoundException(Properties.Resources.DRV_DIR_NOT_RETRIEVE);
            return driverDirectory.ToString();
        }


        delegate bool undoInstall();

        /// <summary>
        /// Installs the port monitor, port,
        /// printer drivers, and VM Print virtual printer
        /// </summary>
        /// <param name="driverSourceDirectory">Directory where the uninstalled printer driver files are located</param>
        /// <param name="driverFilesToCopy">An array containing the printer driver's filenames</param>
        /// <param name="dependentFilesToCopy">An array containing dependent filenames</param>
        /// <returns>true if installation suceeds, false if failed</returns>
        public bool InstallVMPrintPrinter(String driverSourceDirectory,
                                            String outputHandlerCommand,
                                            String outputHandlerArguments)
        {

            bool printerInstalled = false;


            Stack<undoInstall> undoInstallActions = new Stack<undoInstall>();

            String driverDirectory = RetrievePrinterDriverDirectory();
            undoInstallActions.Push(this.DeleteVMPrintPortMonitorDll);
            //仮想プリンターのポートモニタを作成する
            if (AddVMPrintPortMonitor(driverSourceDirectory))
            {
                this.logEventSource.TraceEvent(TraceEventType.Verbose,
                                               (int)TraceEventType.Verbose,
                                               Properties.Resources.MONITOR_INSTALLED);
                undoInstallActions.Push(this.RemoveVMPrintPortMonitor);
                //仮想プリンターのファイルをコピーする
                if (CopyPrinterDriverFiles(driverSourceDirectory, printerDriverFiles.Concat(printerDriverDependentFiles).ToArray()))
                {
                    this.logEventSource.TraceEvent(TraceEventType.Verbose,
                                                   (int)TraceEventType.Verbose,
                                                   Properties.Resources.DRV_ALREADYEXISTS);
                    undoInstallActions.Push(this.RemoveVMPrintPortMonitor);
                    //仮想プリンターのポートを作成する
                    if (AddVMPrintPort())
                    {
                        this.logEventSource.TraceEvent(TraceEventType.Verbose,
                                                       (int)TraceEventType.Verbose,
                                                       Properties.Resources.PORT_ADDED);
                        undoInstallActions.Push(this.RemoveVMPrintPrinterDriver);
                        //仮想プリンターのドライバーをインストールする
                        if (InstallVMPrintPrinterDriver())
                        {
                            this.logEventSource.TraceEvent(TraceEventType.Verbose,
                                                           (int)TraceEventType.Verbose,
                                                           Properties.Resources.DRV_INSTALLED);
                            undoInstallActions.Push(this.DeleteVMPrintPrinter);
                            //仮想プリンターをシステムに追加する
                            if (AddVMPrintPrinter())
                            {
                                this.logEventSource.TraceEvent(TraceEventType.Verbose,
                                                               (int)TraceEventType.Verbose,
                                                               Properties.Resources.VM_PRINT_INSTALLED);
                                undoInstallActions.Push(this.RemoveVMPrintPortConfig);
                                //仮想プリンターを配置する
                                if (ConfigureVMPrintPort(outputHandlerCommand, outputHandlerArguments))
                                {
                                    this.logEventSource.TraceEvent(TraceEventType.Verbose,
                                                                   (int)TraceEventType.Verbose,
                                                                   Properties.Resources.VM_PRINT_CONFIGURED);
                                    printerInstalled = true;
                                }
                                else
                                    // Failed to configure port
                                    this.logEventSource.TraceEvent(TraceEventType.Error,
                                                                    (int)TraceEventType.Error,
                                                                    Properties.Resources.INFO_INSTALLCONFIGPORT_FAILED);
                            }
                            else
                                // Failed to install printer
                                this.logEventSource.TraceEvent(TraceEventType.Error,
                                                                (int)TraceEventType.Error,
                                                                Properties.Resources.INFO_INSTALLPRINTER_FAILED);
                        }
                        else
                            // Failed to install printer driver
                            this.logEventSource.TraceEvent(TraceEventType.Error,
                                                            (int)TraceEventType.Error,
                                                            Properties.Resources.INFO_INSTALLPRINTERDRIVER_FAILED);
                    }
                    else
                        // Failed to add printer port
                        this.logEventSource.TraceEvent(TraceEventType.Error,
                                                        (int)TraceEventType.Error,
                                                        Properties.Resources.INFO_INSTALLPORT_FAILED);
                }
                else
                    //Failed to copy printer driver files
                    this.logEventSource.TraceEvent(TraceEventType.Error,
                                                    (int)TraceEventType.Error,
                                                    Properties.Resources.INFO_INSTALLCOPYDRIVER_FAILED);
            }
            else
                //Failed to add port monitor
                this.logEventSource.TraceEvent(TraceEventType.Error,
                                                (int)TraceEventType.Error,
                                                Properties.Resources.INFO_INSTALLPORTMONITOR_FAILED);

            //インストール失敗の場合、アンインストールする
            if (printerInstalled == false)
            {
                // Printer installation failed -
                // undo all the install steps
                while (undoInstallActions.Count > 0)
                {
                    undoInstall undoAction = undoInstallActions.Pop();
                    try
                    {
                        if (!undoAction())
                        {
                            this.logEventSource.TraceEvent(TraceEventType.Error,
                                                            (int)TraceEventType.Error,
                                                            String.Format(Properties.Resources.INSTALL_ROLLBACK_FAILURE_AT_FUNCTION, undoAction.Method.Name));
                        }
                    }
                    catch (Win32Exception win32Ex)
                    {
                        this.logEventSource.TraceEvent(TraceEventType.Error,
                                                        (int)TraceEventType.Error,
                                                        String.Format(Properties.Resources.INSTALL_ROLLBACK_FAILURE_AT_FUNCTION, undoAction.Method.Name) +
                                                        String.Format(Properties.Resources.WIN32ERROR, win32Ex.ErrorCode.ToString()));
                    }
                }
            }
            this.logEventSource.Flush();
            return printerInstalled;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool UninstallVMPrintPrinter()
        {
            bool printerUninstalledCleanly = true;

            //仮想プリンターを削除する
            if (!DeleteVMPrintPrinter())
                printerUninstalledCleanly = false;
            //仮想プリンターのドライバーを削除する
            if (!RemoveVMPrintPrinterDriver())
                printerUninstalledCleanly = false;
            //仮想プリンターのポートを削除する
            if (!DeleteVMPrintPort())
                printerUninstalledCleanly = false;
            //仮想プリンターのポートモニターを削除する
            if (!RemoveVMPrintPortMonitor())
                printerUninstalledCleanly = false;
            //仮想プリンターの配置情報を削除する
            if (!RemoveVMPrintPortConfig())
                printerUninstalledCleanly = false;
            DeleteVMPrintPortMonitorDll();
            return printerUninstalledCleanly;
        }

        private bool CopyPrinterDriverFiles(String driverSourceDirectory,
                                            String[] filesToCopy)
        {
            bool filesCopied = false;
            String driverDestinationDirectory = RetrievePrinterDriverDirectory();
            try
            {
                for (int loop = 0; loop < filesToCopy.Length; loop++)
                {
                    String fileSourcePath = Path.Combine(driverSourceDirectory, filesToCopy[loop]);
                    String fileDestinationPath = Path.Combine(driverDestinationDirectory, filesToCopy[loop]);
                    try
                    {
                        File.Copy(fileSourcePath, fileDestinationPath);
                    }
                    catch (PathTooLongException)
                    {
                        // Will be caught by outer
                        // IOException catch block
                        throw;
                    }
                    catch (DirectoryNotFoundException)
                    {
                        // Will be caught by outer
                        // IOException catch block
                        throw;
                    }
                    catch (FileNotFoundException)
                    {
                        // Will be caught by outer
                        // IOException catch block
                        throw;
                    }
                    catch (IOException)
                    {
                        // Just keep going - file was already there
                        // Not really a problem
                        logEventSource.TraceEvent(TraceEventType.Verbose,
                                                  (int)TraceEventType.Verbose,
                                                  String.Format(Properties.Resources.FILENOTCOPIED_ALREADYEXISTS, fileDestinationPath));
                        continue;
                    }
                }
                filesCopied = true;
            }
            catch (IOException ioEx)
            { 
                logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          String.Format(Properties.Resources.FILENOTCOPIED_PRINTERDRIVER, ioEx.Message));
            }
            catch (UnauthorizedAccessException unauthorizedEx)
            {
                logEventSource.TraceEvent(TraceEventType.Error,
                            (int)TraceEventType.Error,
                            String.Format(Properties.Resources.FILENOTCOPIED_PRINTERDRIVER, unauthorizedEx.Message));
            }
            catch (NotSupportedException notSupportedEx)
            {
                logEventSource.TraceEvent(TraceEventType.Error,
                    (int)TraceEventType.Error,
                    String.Format(Properties.Resources.FILENOTCOPIED_PRINTERDRIVER, notSupportedEx.Message));
            }


            return filesCopied;
        }

        private bool DeletePrinterDriverFiles(String driverSourceDirectory,
                                              String[] filesToDelete)
        {
            bool allFilesDeleted = true;
            for (int loop = 0; loop < filesToDelete.Length; loop++)
            {
                try
                {
                    File.Delete(Path.Combine(driverSourceDirectory, filesToDelete[loop]));
                }
                catch
                {
                    allFilesDeleted = false;
                }
            }
            return allFilesDeleted;
        }


#if DEBUG
        public bool IsPrinterDriverInstalled_Test(String driverName)
        {
            return IsPrinterDriverInstalled(driverName);
        }
#endif
        private bool IsPrinterDriverInstalled(String driverName)
        {
            bool driverInstalled = false;
            List<DRIVER_INFO_6> installedDrivers = EnumeratePrinterDrivers();
            foreach (DRIVER_INFO_6 printerDriver in installedDrivers)
            {
                //プリンターの名前が一致する場合、Trueを戻る
                if (printerDriver.pName == driverName)
                {
                    driverInstalled = true;
                    break;
                }
            }
            return driverInstalled;
        }

        public List<DRIVER_INFO_6> EnumeratePrinterDrivers()
        {
            List<DRIVER_INFO_6> installedPrinterDrivers = new List<DRIVER_INFO_6>();

            uint pcbNeeded = 0;
            uint pcReturned = 0;

            //プリンターのドライバーを取得する
            if (!NativeMethods.EnumPrinterDrivers(null, ENVIRONMENT_64, 6, IntPtr.Zero, 0, ref pcbNeeded, ref pcReturned))
            {
                IntPtr pDrivers = Marshal.AllocHGlobal((int)pcbNeeded);
                if (NativeMethods.EnumPrinterDrivers(null, ENVIRONMENT_64, 6, pDrivers, pcbNeeded, ref pcbNeeded, ref pcReturned))
                {
                    IntPtr currentDriver = pDrivers;
                    for (int loop = 0; loop < pcReturned; loop++)
                    {
                        installedPrinterDrivers.Add((DRIVER_INFO_6)Marshal.PtrToStructure(currentDriver, typeof(DRIVER_INFO_6)));
                        //currentDriver = (IntPtr)(currentDriver.ToInt32() + Marshal.SizeOf(typeof(DRIVER_INFO_6)));
                        currentDriver = IntPtr.Add(currentDriver, Marshal.SizeOf(typeof(DRIVER_INFO_6)));
                    }
                    Marshal.FreeHGlobal(pDrivers);
                }
                else
                {
                    // Failed to enumerate printer drivers
                    throw new Win32Exception(Marshal.GetLastWin32Error(), Properties.Resources.DRV_NOT_ENUMERATE);
                }
            }
            else
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), Properties.Resources.ENUMDRV_ZERO_SIZE_BUFFER);
            }

            return installedPrinterDrivers;
        }

        private bool InstallVMPrintPrinterDriver()
        {
            bool vmPrintPrinterDriverInstalled = false;

            //ドライバーが未インストールの場合、インストールする
            if (!IsPrinterDriverInstalled(Properties.Resources.DRIVER_NAME))
            {
                String driverSourceDirectory = RetrievePrinterDriverDirectory();

                StringBuilder nullTerminatedDependentFiles = new StringBuilder();
                if (printerDriverDependentFiles.Length > 0)
                {
                    for (int loop = 0; loop <= printerDriverDependentFiles.GetUpperBound(0); loop++)
                    {
                        nullTerminatedDependentFiles.Append(printerDriverDependentFiles[loop]);
                        nullTerminatedDependentFiles.Append("\0");
                    }
                    nullTerminatedDependentFiles.Append("\0");
                }
                else
                {
                    nullTerminatedDependentFiles.Append("\0\0");
                }

                DRIVER_INFO_6 printerDriverInfo = new DRIVER_INFO_6();

                printerDriverInfo.cVersion = 3;
                printerDriverInfo.pName = Properties.Resources.DRIVER_NAME;
                printerDriverInfo.pEnvironment = ENVIRONMENT_64;
                printerDriverInfo.pDriverPath = Path.Combine(driverSourceDirectory, DRIVERFILE);
                printerDriverInfo.pConfigFile = Path.Combine(driverSourceDirectory, DRIVERUIFILE);
                printerDriverInfo.pHelpFile = Path.Combine(driverSourceDirectory, DRIVERHELPFILE);
                printerDriverInfo.pDataFile = Path.Combine(driverSourceDirectory, DRIVERDATAFILE);
                printerDriverInfo.pDependentFiles = nullTerminatedDependentFiles.ToString();

                printerDriverInfo.pMonitorName = PORTMONITOR;
                printerDriverInfo.pDefaultDataType = String.Empty;
                printerDriverInfo.dwlDriverVersion = 0x0000000200000000U;
                printerDriverInfo.pszMfgName = Properties.Resources.DRIVER_MANUFACTURER;
                printerDriverInfo.pszHardwareID = Properties.Resources.HARDWARE_ID;
                printerDriverInfo.pszProvider = Properties.Resources.DRIVER_MANUFACTURER;


                vmPrintPrinterDriverInstalled = InstallPrinterDriver(ref printerDriverInfo);
            }
            else
            {
                vmPrintPrinterDriverInstalled = true; // Driver is already installed, we'll just use the installed driver
            }

            return vmPrintPrinterDriverInstalled;
        }

        private bool InstallPrinterDriver(ref DRIVER_INFO_6 printerDriverInfo)
        {
            bool printerDriverInstalled = false;

            printerDriverInstalled = NativeMethods.AddPrinterDriver(null, 6, ref printerDriverInfo);
            if (printerDriverInstalled == false)
            {
                //int lastWinError = Marshal.GetLastWin32Error();
                //throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not add printer VM Print printer driver.");
                logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          Properties.Resources.VM_DRV_NOT_ADDED +
                                          String.Format(Properties.Resources.WIN32ERROR, Marshal.GetLastWin32Error().ToString()));
            }
            return printerDriverInstalled;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool RemoveVMPrintPrinterDriver()
        {
            //ドライバーを削除する
            bool driverRemoved = NativeMethods.DeletePrinterDriverEx(null, ENVIRONMENT_64, Properties.Resources.DRIVER_NAME, DPD_DELETE_UNUSED_FILES, 3);
            if (!driverRemoved)
            {
                logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          Properties.Resources.VM_DRV_NOT_REMOVE +
                                          String.Format(Properties.Resources.WIN32ERROR, Marshal.GetLastWin32Error().ToString()));
            }
            return driverRemoved;
        }


        private bool AddVMPrintPrinter()
        {
            //ﾌﾟﾘﾝﾀｰを追加する
            bool printerAdded = false;
            PRINTER_INFO_2 vmPrintPrinter = new PRINTER_INFO_2();

            vmPrintPrinter.pServerName = null;
            vmPrintPrinter.pPrinterName = Properties.Resources.PRINTER_NAME;
            vmPrintPrinter.pPortName = PORTNAME;
            vmPrintPrinter.pDriverName = Properties.Resources.DRIVER_NAME;
            vmPrintPrinter.pPrintProcessor = Properties.Resources.PRINT_PROCESOR;
            vmPrintPrinter.pDatatype = "RAW";
            vmPrintPrinter.Attributes = 0x00000002;
            vmPrintPrinter.Attributes = PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST;

            int vmPrintPrinterHandle = NativeMethods.AddPrinter(null, 2, ref vmPrintPrinter);
            if (vmPrintPrinterHandle != 0)
            {
                // Added ok
                int closeCode = NativeMethods.ClosePrinter((IntPtr)vmPrintPrinterHandle);
                printerAdded = true;
            }
            else
            {
                logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          Properties.Resources.VM_NOT_ADDED + 
                                          String.Format(Properties.Resources.WIN32ERROR, Marshal.GetLastWin32Error().ToString()));
            }
            return printerAdded;
        }

        private bool DeleteVMPrintPrinter()
        {
            //ﾌﾟﾘﾝﾀｰを削除する
            bool printerDeleted = false;

            PRINTER_DEFAULTS vmPrintDefaults = new PRINTER_DEFAULTS();
            vmPrintDefaults.DesiredAccess = 0x000F000C; // All access
            vmPrintDefaults.pDatatype = null;
            vmPrintDefaults.pDevMode = IntPtr.Zero;

            IntPtr vmPrintHandle = IntPtr.Zero;
            try
            {
                //ﾌﾟﾘﾝﾀｰのハンドルを取得できる場合、該当プリンターを削除する
                if (NativeMethods.OpenPrinter(Properties.Resources.PRINTER_NAME, ref vmPrintHandle, vmPrintDefaults) != 0)
                {
                    if (NativeMethods.DeletePrinter(vmPrintHandle))
                        printerDeleted = true;
                }
                else
                {
                    logEventSource.TraceEvent(TraceEventType.Error,
                                              (int)TraceEventType.Error,
                                              Properties.Resources.VM_NOT_DELETE +
                                              String.Format(Properties.Resources.WIN32ERROR, Marshal.GetLastWin32Error().ToString()));
                }
            }
            finally
            {
                if (vmPrintHandle != IntPtr.Zero) NativeMethods.ClosePrinter(vmPrintHandle);
            }
            return printerDeleted;
        }


        public bool IsVMPrintPrinterInstalled()
        {
            bool vmPrintInstalled = false;

            PRINTER_DEFAULTS vmPrintDefaults = new PRINTER_DEFAULTS();
            vmPrintDefaults.DesiredAccess = 0x00008; // Use access
            vmPrintDefaults.pDatatype = null;
            vmPrintDefaults.pDevMode = IntPtr.Zero;

            IntPtr vmPrintHandle = IntPtr.Zero;
            //ﾌﾟﾘﾝﾀｰが開ける場合、インストール済とする
            if (NativeMethods.OpenPrinter(Properties.Resources.PRINTER_NAME, ref vmPrintHandle, vmPrintDefaults) != 0)
            {
                vmPrintInstalled = true;
            }
            else
            {
                int errorCode = Marshal.GetLastWin32Error();
                if (errorCode == 0x5) vmPrintInstalled = true; // Printer is installed, but user
                                                                 // has no privileges to use it
            }

            return vmPrintInstalled;
        }

        #endregion





        #region Configuration and Registry changes

#if DEBUG
        public bool ConfigureVMPrintPort_Test()
        {
            return ConfigureVMPrintPort();
        }
#endif

        private bool ConfigureVMPrintPort()
        {
            return ConfigureVMPrintPort(String.Empty, String.Empty);

        }

        
        private bool ConfigureVMPrintPort(String commandValue,
                                            String argumentsValue)
        {
            bool registryChangesMade = false;
            // Add all the registry info
            // for the port and monitor
            RegistryKey portConfiguration;
            try
            {
                portConfiguration = Registry.LocalMachine.CreateSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors\\" +  PORTMONITOR + "\\Ports\\" + PORTNAME);
                portConfiguration.SetValue("Description", Properties.Resources.PORT_DESCRIPTION, RegistryValueKind.String);
                portConfiguration.SetValue("Command", commandValue, RegistryValueKind.String);
                portConfiguration.SetValue("Arguments", argumentsValue, RegistryValueKind.String);
                portConfiguration.SetValue("Printer", Properties.Resources.PRINTER_NAME, RegistryValueKind.String);
                portConfiguration.SetValue("Output", 0, RegistryValueKind.DWord);
                portConfiguration.SetValue("ShowWindow", 2, RegistryValueKind.DWord);
                portConfiguration.SetValue("RunUser", 1, RegistryValueKind.DWord);
                portConfiguration.SetValue("Delay", 300, RegistryValueKind.DWord);
                portConfiguration.SetValue("LogFileUse", 0, RegistryValueKind.DWord);
                portConfiguration.SetValue("LogFileName", "", RegistryValueKind.String);
                portConfiguration.SetValue("LogFileDebug", 0, RegistryValueKind.DWord);
                portConfiguration.SetValue("PrintError", 0, RegistryValueKind.DWord);
                registryChangesMade = true;
            }

            catch (UnauthorizedAccessException unauthorizedEx)
            {
                logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          String.Format(Properties.Resources.REGISTRYCONFIG_NOT_ADDED, unauthorizedEx.Message));
            }
            catch (SecurityException securityEx)
            {
                logEventSource.TraceEvent(TraceEventType.Error,
                            (int)TraceEventType.Error,
                            String.Format(Properties.Resources.REGISTRYCONFIG_NOT_ADDED, securityEx.Message));
            }

            return registryChangesMade;
        }

        private bool RemoveVMPrintPortConfig()
        {
            bool registryEntriesRemoved = false;

            try
            {
                Registry.LocalMachine.DeleteSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors\\" + PORTMONITOR + "\\Ports\\" + PORTNAME, false);
                registryEntriesRemoved = true;
            }
            catch (UnauthorizedAccessException unauthorizedEx)
            {
                logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          String.Format(Properties.Resources.REGISTRYCONFIG_NOT_DELETED, unauthorizedEx.Message));
            }

            return registryEntriesRemoved;

        }

        #endregion

    }
}
