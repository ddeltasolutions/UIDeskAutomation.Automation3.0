using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace UIDeskAutomationLib
{
    internal partial class FrmDummy : Form
    {
        public FrmDummy()
        {
            InitializeComponent();

            timer.AutoReset = false;

            timer.Elapsed += timer_Elapsed;
        }

        private void FrmDummy_Load(object sender, EventArgs e)
        {
            //SetTimer();
            
        }

        private void SetTimer()
        {
            if (timer.Enabled == true)
            {
                return;
            }
            
            //timer.AutoReset = false;

            //timer.Elapsed += timer_Elapsed;

            try
            {
                timer.Start();
            }
            catch { }
        }

        void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                this.Close();

                //this.Visible = false;
            }
            catch { }
        }

        /*private void FrmDummy_Paint(object sender, PaintEventArgs e)
        {
            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();

            Rectangle rectLeft = new Rectangle(0, 0, 5, this.Height);

            Rectangle rectTop = new Rectangle(0, 0, this.Width, 5);

            

            path.AddRectangles(new Rectangle[] { rectLeft, rectTop });

            //e.Graphics.DrawPath(Pens.Red, path);

            this.Region = new System.Drawing.Region(path);
        }*/

        System.Timers.Timer timer = new System.Timers.Timer(200);
    }
}
