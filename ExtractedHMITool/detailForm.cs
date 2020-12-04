using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExtractedHMITool
{
    public partial class detailForm : Form
    {
        string path = "";

        public detailForm()
        {
            InitializeComponent();
        }
        
        public string path_RTF
        {
            get
            {
                return path;
            }
            set
            {
                path = value;
                if (path.Equals(""))
                {

                    animationTextBox.Clear();
                }
                else
                {

                    animationTextBox.LoadFile(@path);

                }
            }
        }
        private void detailForm_Load(object sender, EventArgs e)
        {


        }

        private void detailForm_Activated(object sender, EventArgs e)
        {

        }
    }
}
