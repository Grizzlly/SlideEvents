using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SlideEvents
{
    public partial class EditEventForm : Form
    {
        public EditEventForm()
        {
            InitializeComponent();
        }

        public EditEventForm(string url)
        {
            InitializeComponent();
            urlBox.Text = url;
        }

        public string GetUrl()
        {
            return urlBox.Text;
        }

    }
}
