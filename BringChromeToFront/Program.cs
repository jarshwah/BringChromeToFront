using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace BringChromeToFront
{
    class Program
    {
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);

        static IEnumerable<IntPtr> EnumerateProcessWindowHandles(int processId)
        {
            var handles = new List<IntPtr>();
            foreach (ProcessThread thread in Process.GetProcessById(processId).Threads)
            {
                EnumThreadWindows(thread.Id,
                    (hWnd, lParam) =>
                    {
                        handles.Add(hWnd);
                        return true;
                    },
                    IntPtr.Zero
                );
            }
            return handles;
        }

        private const int SW_RESTORE = 9; // maximise from minimised

        static void Main(string[] args)
        {
            // Test Cases:
            //
            // 1. Minimise chrome with the correct tab already selected:
            //    Failure: finds the tab text, but doesn't maximise the browser
            //    Reason: SetForegroundWindow works, but the IsIconic() check is negative due to incorrect window handle
            //
            // 2. Minimise chrome with incorrect tab selected:
            //    Pass: we reach the correct window handle after trying multiple incorrect window handles
            //
            // 3. Maximise chrome with correct tab selected
            //    Pass: SetForegroundWindow brings the browser to front. IsIconic doesn't come into play
            //
            // 4. Have two browser windows open (to ensure we're iterating all windows)
            //    Pass: As long as the first browser isn't minimised with the correct tab already selected
            

            const string process = "chrome"; // process name we're looking for
            const string tabText = "hacker"; // the window/tab text to find
            const int delay = 50; // ms delay after sending key presses to window - 100 seems to be standard, but 50 still gives me the right results
            // run under a different thread, because we don't want to delay the UI thread of the application
            ThreadPool.QueueUserWorkItem(
                _ =>
                {
                    BringToFront(process, tabText, delay);
                    Console.WriteLine("Finished - press ENTER key to exit");
                }
            );
            Console.ReadKey();
        }

        static void BringToFront(string process, string title, int delay)
        {
            var procs = Process.GetProcessesByName(process);
            if (procs.Length == 0)
            {
                Console.WriteLine("Process {0} not running. Can't bring to front.", process);
                return;
            }

            // we have to iterate through all, as there may be multiple windows open
            foreach (var proc in procs)
            {
                // only inspect processes with a valid window handle
                Console.WriteLine("Examining ProcessId {0}", proc.Id);
                foreach (var handle in EnumerateProcessWindowHandles(proc.Id))
                {
                    Console.WriteLine(
                        "Examining window handle {0} for ProcessId {1}",
                        handle.ToInt64(),
                        proc.Id);

                    if (handle.ToInt64() > 0)
                    {
                        Console.WriteLine("Window with ProcessId {0} has a valid hWnd {1}",
                            proc.Id, handle.ToInt64());

                        if (IsIconic(handle))
                        {
                            Console.WriteLine("Window is minimized -> maximising now");
                            ShowWindow(handle, SW_RESTORE);
                            Thread.Sleep(delay);
                        }

                        SetForegroundWindow(handle);
                        if (Process.GetProcessById(proc.Id).MainWindowTitle.ToLower().Contains(title.ToLower()))
                        {
                            Console.WriteLine("Found tab <{0}> early - no need to examine other procs", title);
                            return;
                        }
                        SendKeys.SendWait("^1"); // first tab
                        Thread.Sleep(delay);

                        // this will be more than the number of actual tabs, because it will include
                        // tabs from all windows, plus things like plugin manager and others, but it
                        // is the best estimation I can come up with.
                        var notChanged = 0;
                        var lastTitle = "";
                        for (var loop = 0; loop < procs.Length; loop++)
                        {
                            // get the process again because the title will have changed by tabbing
                            var currentTitle = Process.GetProcessById(proc.Id).MainWindowTitle;
                            Console.WriteLine("Checking Window <{0}>", currentTitle);
                            
                            if (string.IsNullOrEmpty(currentTitle))
                            {
                                break; // looking at an internal window
                            }

                            if (currentTitle.ToLower().Contains(title.ToLower()))
                            {
                                Console.WriteLine("Found tab <{0}>", title);
                                return;
                            }

                            // optimisation
                            if (lastTitle == currentTitle)
                            {
                                notChanged++;
                                if (notChanged >= 4)
                                // if the user has 4 of the same tabs open in a row, this will incorrectly skip.
                                {
                                    Console.WriteLine(
                                        "Optimisation: We've seen this title 4 times in a row. Assume wrong window and skip.");
                                    break;
                                }
                            }
                            else
                            {
                                notChanged = 0;
                            }
                            lastTitle = currentTitle;

                            SendKeys.SendWait("^{TAB}"); // cycle through tabs
                            Thread.Sleep(delay);
                        }
                    }
                }
            }

            Console.WriteLine("Could not identify the correct tab to Force to front.");
        }
    }
}
