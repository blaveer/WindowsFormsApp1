using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Program_using
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            //是否显示cmd
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            //重定向输入流
            process.StartInfo.RedirectStandardInput = true;
            //重定向输出流
            process.StartInfo.RedirectStandardOutput = true;
            string strCmd = "ping " + textBox2.Text.ToString();
            process.Start();
            process.StandardInput.WriteLine(strCmd);
            process.StandardInput.WriteLine("exit");
            //获取输出信息
            textBox1.Text = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            process.Close();
        }


        private void StartCmd()
        {
            textBox1.Text = "";
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            // 是否在新窗口中启动该进程的值   
            process.StartInfo.CreateNoWindow = true;
            // 重定向输入流  
            process.StartInfo.RedirectStandardInput = true;
            // 重定向输出流
            process.StartInfo.RedirectStandardOutput = true;
            //使ping命令执行
            string strCmd = "ping " + textBox2.Text;
            process.Start();
            process.StandardInput.WriteLine(strCmd);
            process.StandardInput.WriteLine("exit");
            process.OutputDataReceived += new DataReceivedEventHandler(strOutputHandler);
            process.BeginOutputReadLine();
        }


        private void strOutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            textBox1.Text += outLine.Data + "\r\n";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            StartCmd();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string physicalAddress = "";
            NetworkInterface[] nice = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface adaper in nice)
            {
                if (adaper.Description == "en0")
                {
                    physicalAddress = adaper.GetPhysicalAddress().ToString();
                    break;
                }
                else
                {
                    physicalAddress = adaper.GetPhysicalAddress().ToString();
                    if (physicalAddress != "")
                    {
                        break;
                    };
                }
            }
            textBox1.Text = physicalAddress;
        }

        private void button4_Click(object sender, EventArgs e)
        {

            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.CreateNoWindow = false;
            //process.StartInfo.UseShellExecute = false;
            //string strCmd = "ping " + textBox2.Text.ToString();
            //process.StartInfo.RedirectStandardInput = true;
            process.Start();
            //process.StandardInput.WriteLine(strCmd);
            //process.StandardInput.WriteLine("exit");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Process[] processes = Process.GetProcessesByName("cmd");
            foreach (Process process in processes)
            {
                process.Kill();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
