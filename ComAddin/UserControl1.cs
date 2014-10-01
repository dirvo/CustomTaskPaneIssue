using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ComAddin
{
    [ComVisible(true)]
    [ProgId("ComAddin.UserControl")]
    [Guid("26B0BF2A-D3DF-4222-A13D-63FD0914945E")]
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }
    }
}
