using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;

namespace Report_generator
{
    public partial class LoadForm : Form
    {
        static LoadForm loadForm = null;
        static Thread loadFormThread = null;
        private const int TIMER_INTERVAL = 250;
        private delegate void CloseDelegate();
        private System.Windows.Forms.Timer timer;
        public LoadForm()
        { 
            InitializeComponent();
            
            timer = new System.Windows.Forms.Timer();
            timer.Interval = TIMER_INTERVAL;
            timer.Enabled = true;
            timer.Start();
        }
        private void LoadForm_Load(object sender, EventArgs e) { timer.Tick += new EventHandler(timer_Tick); }

        void timer_Tick(object sender, EventArgs e)
        {
            switch (this.dynamicLabel1.Text)
            {
                case "--": this.dynamicLabel1.Text = @" \"; this.dynamicLabel2.Text = @" \"; break;
                case @" \": this.dynamicLabel1.Text = @" |"; this.dynamicLabel2.Text = @" |"; break;
                case @" |": this.dynamicLabel1.Text = @" /"; this.dynamicLabel2.Text = @" /"; break;
                case @" /": this.dynamicLabel1.Text = @"--"; this.dynamicLabel2.Text = @"--"; break;
            }
        }
        static private void ShowForm()
        {
            loadForm = new LoadForm();
            loadForm.ShowDialog();
            //Application.Run(loadForm);
        }
        // A static method to close the SplashScreen
        static public void CloseForm()
        {
            //loadForm.Close();
            //loadFormThread = null;
            //loadForm = null;
            loadForm.Invoke(new CloseDelegate(LoadForm.CloseFormInternal));
        }
        static private void CloseFormInternal()
        { loadForm.Close(); loadForm = null; loadFormThread = null; }
        static public void ShowLoadForm()
        {
            // Make sure it is only launched once.
            if (loadForm != null)
                return;
            loadFormThread = new Thread(new ThreadStart(LoadForm.ShowForm));
            loadFormThread.IsBackground = true;
            loadFormThread.SetApartmentState(ApartmentState.STA);
            loadFormThread.Start();
            //while (loadForm == null || loadForm.IsHandleCreated == false) { System.Threading.Thread.Sleep(TIMER_INTERVAL); }
        }

    }
}
