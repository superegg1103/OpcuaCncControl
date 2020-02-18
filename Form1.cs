using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PickPosition
{
    public partial class Form1 : Form
    {
        static Thread PickThread = new Thread(PickPosition.Picking);
        public static string localPath;
        public static string textMessage = "Connecting...";
        public static string localIP;
        public static string lasorIP;

        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if(PickThread.IsAlive == false)
            {
                localIP = textBox3.Text;
                lasorIP = textBox4.Text;
                timer1.Enabled = true;
                localPath = textBox1.Text;
                PickPosition.flagflag = 1;
                PickThread.Start();
            }
        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            PickPosition.flagflag = 0;
            PickThread.Interrupt();
            textMessage += "\r\nFinish";
            ResetThread();
        }
        
        private void Timer1_Tick(object sender, EventArgs e)
        {
            textBox2.Text = textMessage;
            textBox2.SelectionStart = this.textBox2.TextLength;
            textBox2.ScrollToCaret();
        }

        public static void ResetThread()
        {
            PickThread = new Thread(PickPosition.Picking);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
