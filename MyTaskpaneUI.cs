using CADBooster.SolidDna;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SolidDNA.Memo.AddIn
{
    /// <summary>
    /// The host control that is registered for COM so that SolidWorks can create an instance of this UI and 
    /// inject it into the taskpane
    /// </summary>
    [ProgId(MyProgId)]
    public partial class MyTaskpaneUI : UserControl, ITaskpaneControl
    {
        #region Private Members

        /// <summary>
        /// Our unique ProgId for SolidWorks to find and load us
        /// </summary>
        private const string MyProgId = "SolidDNA.Memo.AddIn.Taskpane";

        #endregion

        #region Public Properties

        /// <summary>
        /// The ProgId for the COM exposure of this UI
        /// </summary>
        public string ProgId => MyProgId;

        #endregion
    }
}
