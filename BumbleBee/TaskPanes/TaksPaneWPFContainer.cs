using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using WPFC = System.Windows.Controls;

namespace BumbleBee.TaskPanes
{
    public partial class TaksPaneWPFContainer<T> : UserControl where T : WPFC.UserControl
    {
        public T Child { get; private set; }

        public TaksPaneWPFContainer(T child)
        {
            InitializeComponent();
            Child = child;
            WpfElementHost.HostContainer.Children.Add(Child);
        }

        public ElementHost WpfElementHost
        {
            get { return wpfElementHost; }
        }
    }
}
