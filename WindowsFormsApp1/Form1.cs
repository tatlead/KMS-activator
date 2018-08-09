using System;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.Principal;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        int g_iWindowsActivateState = -1;

        public Form1()
        {
            //Check admin
            if (!IsAdministrator())
            {
                var exeName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                ProcessStartInfo startInfo = new ProcessStartInfo(exeName);
                startInfo.Verb = "runas";
                System.Diagnostics.Process.Start(startInfo);
                Environment.Exit(0);
            }

            InitializeComponent();

            label4.Text = "";
            label5.Text = "";
            textBox1.MaxLength = 29;

            string strKMSkey = GetKMSSetUpKey();
            string strOSName = GetOSFriendlyName();

            UpdateWindowsStatus();
            if (g_iWindowsActivateState == 1) //activated permanently
            {
                SetActivatingState();
                label5.Text = strOSName + " already activated permanently";
            } 
            else if (strKMSkey == "") //Cannot find the KMS key
            {
                SetActivatingState();
                label5.Text = "Sorry, " + strOSName + " is not supported for KMS activation.";
            }
            else
            {
                textBox1.Text = strKMSkey;
                label1.Text = strOSName + " detected, KMS Key autofilled.";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetActivatingState();
            label5.Text = "Activating...";

            progressBar1.Value = 0;
            progressBar1.Value += 20;
            WaitNSeconds(1);

            if (GetWindowsActivatedState() == 1)
            {
                label5.Text = GetOSFriendlyName() + " already Activated Parmanently! KMS Activation cancelled.";
                return;
            }
 
            string strKMSkey = textBox1.Text;
            string strKMSser = textBox2.Text;
           
            //install KMS key
            string strKeystatus = ExecuteCommand("nologo slmgr.vbs /ipk " + strKMSkey);
            if (!strKeystatus.Contains("successfully"))
            {
                label5.Text = "Fail to install KMS Key.";
                SetNotActivatingState();
                return;
            }
            label5.Text = "KMS Key installed successfully!";
            progressBar1.Value += 20;
            WaitNSeconds(1);

            //install KMS server
            string strServerstatus = ExecuteCommand("nologo slmgr.vbs /skms " + strKMSser);
            if (!strServerstatus.Contains("successfully"))
            {
                //This area seems really reachable.. slmgr.vbs /skms will return successfully everytime
                label5.Text = "Fail to connect KMS server.";
                SetNotActivatingState();
                return;
            }
            label5.Text = "Set KMS server to " + strKMSser + " successfully!";
            progressBar1.Value += 20;
            WaitNSeconds(1);

            //Activate windows
            string strActivatestatus = ExecuteCommand("nologo slmgr.vbs /ato");
            if (!strActivatestatus.Contains("successfully"))
            {
                label5.Text = "The KMS server is unavailable.";
                SetNotActivatingState();
                return;
            }
            label5.Text = "Connected to " + strKMSser + " successfully!";
            progressBar1.Value += 20;
            WaitNSeconds(1);

            //Check activate state
            int iActivateState = GetWindowsActivatedState();
            if (iActivateState == -1)
            {
                label5.Text = "Fail to activate. Try another KMS server.";
                SetNotActivatingState();
                return;
            }
            label5.Text = "Activate successfully!";
            progressBar1.Value += 20;
            UpdateWindowsStatus();

            WaitNSeconds(2);
            SetNotActivatingState();
        }

        private void SetActivatingState()
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;
            progressBar1.Value = 0;
        }

        private void SetNotActivatingState()
        {
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            progressBar1.Value = 0;
        }

        private void UpdateWindowsStatus()
        {
            g_iWindowsActivateState = GetWindowsActivatedState();
            if (g_iWindowsActivateState >= 0)
            {
                label6.ForeColor = Color.Green;
                if (g_iWindowsActivateState == 1)
                {
                    label6.Text = "Parmanently Activated";
                }
                else label6.Text = "Activated";
            }
            else
            {
                label6.ForeColor = Color.Red;
                label6.Text = "Not Activated";
            }
        }

        //Return output
        private string ExecuteCommand(string strCommand)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = "cscript.exe";
            psi.WorkingDirectory = System.Environment.GetEnvironmentVariable("SystemRoot") + @"\System32";
            psi.UseShellExecute = false;
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            psi.RedirectStandardInput = true;
            psi.RedirectStandardOutput = true;
            psi.CreateNoWindow = true;
            psi.Arguments = @"//" + strCommand;

            Process scriptProc = new Process();
            scriptProc.StartInfo = psi;
            scriptProc.Start();

            return scriptProc.StandardOutput.ReadToEnd();
        }

        //https://docs.microsoft.com/en-us/windows-server/get-started/kmsclientkeys
        private string GetKMSSetUpKey()
        {
            switch (GetOSFriendlyName())
            {
                //Windows 7
                case ("Windows 7 Professional"): return "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4";
                case ("Windows 7 Professional N"): return "MRPKT-YTG23-K7D7T-X2JMM-QY7MG";
                case ("Windows 7 Professional E"): return "W82YF-2Q76Y-63HXB-FGJG9-GF7QX";
                case ("Windows 7 Enterprise"): return "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH";
                case ("Windows 7 Enterprise N"): return "YDRBP-3D83W-TY26F-D46B2-XCKRJ";
                case ("Windows 7 Enterprise E"): return "C29WB-22CC8-VJ326-GHFJW-H9DH4";

                //Windows 8
                case ("Windows 8 Professional"): return "NG4HW-VH26C-733KW-K6F98-J8CK4";
                case ("Windows 8 Professional N"): return "XCVCF-2NXM9-723PB-MHCB7-2RYQQ";
                case ("Windows 8 Enterprise"): return "32JNW-9KQ84-P47T8-D8GGY-CWCK7";
                case ("Windows 8 Enterprise N"): return "JMNMF-RHW7P-DMY6X-RF3DR-X2BQT";

                //Windows 8.1
                case ("Windows 8.1 Professional"): return "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9";
                case ("Windows 8.1 Professional N"): return "HMCNV-VVBFX-7HMBH-CTY9B-B4FXY";
                case ("Windows 8.1 Enterprise"): return "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7";
                case ("Windows 8.1 Enterprise N"): return "TT4HM-HN7YT-62K67-RGRQJ-JFFXW";

                //Windows 10
                case ("Windows 10 Professional"): return "W269N-WFGWX-YVC9B-4J6C9-T83GX";
                case ("Windows 10 Professional N"): return "MH37W-N47XK-V7XM9-C7227-GCQG9";
                case ("Windows 10 Enterprise"): return "NPPR9-FWDCX-D2C8J-H872K-2YT43";
                case ("Windows 10 Enterprise N"): return "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4";
                case ("Windows 10 Education"): return "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2";
                case ("Windows 10 Education N"): return "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ";
                case ("Windows 10 Enterprise 2015 LTSB"): return "WNMTR-4C88C-JK8YV-HQ7T2-76DF9";
                case ("Windows 10 Enterprise 2015 LTSB N"): return "2F77B-TNFGY-69QQF-B8YKP-D69TJ";
                case ("Windows 10 Enterprise 2016 LTSB"): return "DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ";
                case ("Windows 10 Enterprise 2016 LTSB N"): return "QFFDN-GRT3P-VKWWX-X7T3R-8B639";
                case ("Windows 10 Professional Education"): return "6TP4R-GNPTD-KYYHQ-7B7DP-J447Y";
                case ("Windows 10 Professional Education N"): return "YVWGF-BXNMC-HTQYQ-CPQ99-66QFC";
                case ("Windows 10 Professional Workstation"): return "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J";
                case ("Windows 10 Professional Workstation N"): return "9FNHH-K3HBT-3W4TD-6383H-6XYWF";

                //Windows Server, version 1803
                case ("Windows Server Datacenter"): return "2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG";
                case ("Windows Server Standard"): return "PTXN8-JFHJM-4WC78-MPCBR-9W4KR";

                //Windows Server 2016
                case ("Windows Server 2016 Datacenter"): return "CB7KF-BWN84-R7R2Y-793K2-8XDDG";
                case ("Windows Server 2016 Standard"): return "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY";
                case ("Windows Server 2016 Essentials"): return "JCKRF-N37P4-C2D82-9YXRT-4M63B";

                //Windows Server 2012 R2
                case ("Windows Server 2012 R2 Server Standard"): return "D2N9P-3P6X9-2R39C-7RTCD-MDVJX";
                case ("Windows Server 2012 R2 Datacenter"): return "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9";
                case ("Windows Server 2012 R2 Essentials"): return "KNC87-3J2TX-XB4WP-VCPJV-M4FWM";

                //Windows Server 2012
                case ("Windows Server 2012"): return "BN3D2-R7TKB-3YPBD-8DRP2-27GG4";
                case ("Windows Server 2012 N"): return "8N2M2-HWPGY-7PGT9-HGDD8-GVGGY";
                case ("Windows Server 2012 Single Language"): return "2WN2H-YGCQR-KFX6K-CD6TF-84YXQ";
                case ("Windows Server 2012 Country Specific"): return "4K36P-JN4VD-GDC6V-KDT89-DYFKP";
                case ("Windows Server 2012 Server Standard"): return "XC9B7-NBPP2-83J2H-RHMBY-92BT4";
                case ("Windows Server 2012 MultiPoint Standard"): return "HM7DN-YVMH3-46JC3-XYTG7-CYQJJ";
                case ("Windows Server 2012 MultiPoint Premium"): return "XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G";
                case ("Windows Server 2012 Datacenter"): return "48HP8-DN98B-MYWDG-T2DCC-8W83P";

                // Windows Server 2008 R2
                case ("Windows Server 2008 R2 Web"): return "6TPJF-RBVHG-WBW2R-86QPH-6RTM4";
                case ("Windows Server 2008 R2 HPC edition"): return "TT8MH-CG224-D3D7Q-498W2-9QCTX";
                case ("Windows Server 2008 R2 Standard"): return "YC6KT-GKW9T-YTKYR-T4X34-R7VHC";
                case ("Windows Server 2008 R2 Enterprise"): return "489J6-VHDMP-X63PK-3K798-CPX3Y";
                case ("Windows Server 2008 R2 Datacenter"): return "74YFP-3QFB3-KQT8W-PMXWJ-7M648";
                case ("Windows Server 2008 R2 for Itanium-based Systems"): return "GT63C-RJFQ3-4GMB6-BRFB9-CB83V";

                //Windows Server 2008
                case ("Windows Web Server 2008"): return "WYR28-R7TFJ-3X2YQ-YCY4H-M249D";
                case ("Windows Server 2008 Standard"): return "TM24T-X9RMF-VWXK6-X8JC9-BFGM2";
                case ("Windows Server 2008 Standard without Hyper-V"): return "W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ";
                case ("Windows Server 2008 Enterprise"): return "YQGMW-MPWTJ-34KDK-48M3W-X4Q6V";
                case ("Windows Server 2008 Enterprise without Hyper-V"): return "39BXF-X8Q23-P2WWT-38T2F-G3FPG";
                case ("Windows Server 2008 HPC"): return "RCTX3-KWVHP-BR6TB-RB6DM-6X7HP";
                case ("Windows Server 2008 Datacenter"): return "7M67G-PC374-GR742-YH8V4-TCBY3";
                case ("Windows Server 2008 Datacenter without Hyper-V"): return "22XQ2-VRXRG-P8D42-K34TD-G3QQC";
                case ("Windows Server 2008 for Itanium-Based Systems"): return "4DWFP-JF3DJ-B7DTH-78FJB-PDRHK";

                default: return "";
            }
        }

        //https://stackoverflow.com/questions/31885302/how-can-i-detect-if-my-app-is-running-on-windows-10
        private string GetOSFriendlyName()
        {
            string subKey = @"SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion";
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey skey = key.OpenSubKey(subKey);
            return skey.GetValue("ProductName").ToString();
        }

        /** Get Windows Activated State
         * return
         * {
         *     -1 = Not Activated
         *      0 = Activated (KMS)
         *      1 = Activated permanently
         * }
        **/
        private int GetWindowsActivatedState()
        {
            //View the License Expiration Date
            string strWindowsStatus = ExecuteCommand("nologo slmgr.vbs /xpr");
            if (strWindowsStatus.Contains("permanently"))
            {
                return 1;
            }
            else if (strWindowsStatus.Contains("expire")) //KMS Activate
            {
                return 0;
            }
            return -1;
        }

        //https://stackoverflow.com/questions/22158278/wait-some-seconds-without-blocking-ui-execution
        private void WaitNSeconds(int segundos)
        {
            if (segundos < 1) return;
            DateTime _desired = DateTime.Now.AddSeconds(segundos);
            while (DateTime.Now < _desired)
            {
                System.Windows.Forms.Application.DoEvents();
            }
        }

        //https://stackoverflow.com/questions/133379/elevating-process-privilege-programmatically
        private static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private bool IsValidKMSkey(string strKMSkey)
        {
            if (strKMSkey.Length == 29)
            {
                return true;
            }
            return false;
        }

        //ping the server
        private bool IsValidKMSser(string strKMSser)
        {
            ProcessStartInfo psi = new ProcessStartInfo("cmd.exe");
            psi.UseShellExecute = false;
            psi.Arguments = "/c ping " + strKMSser;
            psi.RedirectStandardInput = true;
            psi.RedirectStandardOutput = true;
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            psi.CreateNoWindow = true;

            Process scriptProc = new Process();
            scriptProc.StartInfo = psi;
            scriptProc.Start();

            string result = scriptProc.StandardOutput.ReadToEnd();
            if (result.Contains("Reply from"))
            {
                return true;
            }
            return false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        //Input KMS key
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        //Input KMS server
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        //Check The KMS key Valid
        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.ReadOnly = true;
            label1.Text = "Checking...";
            WaitNSeconds(1);

            string strKMSkey = textBox1.Text;
            if (IsValidKMSkey(strKMSkey))
            {
                string strKeystatus = ExecuteCommand("nologo slmgr.vbs /ipk " + strKMSkey);
                if (strKeystatus.Contains("successfully"))
                {
                    label1.Text = "The KMS Key is VALID.";
                }
                else label1.Text = "The KMS Key is INVALID.";
            }
            else label1.Text = "The KMS Key is INVALID.";

            textBox1.ReadOnly = false;
        }

        //Check The KMS server reachable
        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.ReadOnly = true;
            label4.Text = "Checking...";
            WaitNSeconds(1);

            string strKMSser = textBox2.Text;
            if (IsValidKMSser(strKMSser))
            {
                label4.Text = "The KMS server is REACHABLE."; 
            }
            else label4.Text = "The KMS server is UNREACHABLE.";

            textBox2.ReadOnly = false;
        }

        //Reminder1
        private void label1_Click(object sender, EventArgs e)
        {

        }

        //Reminder2
        private void label4_Click(object sender, EventArgs e)
        {

        }

        //Large Label
        private void label6_Click(object sender, EventArgs e)
        {

        }

        //Sourcecode
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel1.LinkVisited = true;
            System.Diagnostics.Process.Start("https://github.com/BattlefieldDuck/KMS-activator");
        }
    }

    internal class ManagementObjectSearcher
    {
    }
}
