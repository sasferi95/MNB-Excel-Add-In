using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MNB_Excel_Add_In
{
    public partial class LogWindow : Form
    {
        public LogWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// after windows is loaded fill the datagridwiev with data
        /// </summary>
        /// <param name="sender">Winodw</param>
        /// <param name="e">eventargs</param>
        private void LogWindow_Load(object sender, EventArgs e)
        {
            try
            {
                this.mNBButtonLogsTableAdapter.Fill(this.excelButtonDataSet.MNBButtonLogs);
            }
            catch(Exception exc)
            {
                MessageBox.Show("Error happened whilegetting logs from database.\n" + exc.Message);
            }
        }

        /// <summary>
        /// when the save button is clicked updates the database based on the changes 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveBtn_Click(object sender, EventArgs e)
        {
            /*when building the application it will use the database in the debug folder's resource folder that's
            why if you restart the program it will copy the original db over the possibly changed database and seems like the changes are lost*/

            try
            {
                this.mNBButtonLogsTableAdapter.Update(this.excelButtonDataSet.MNBButtonLogs);
            }
            catch(Exception exc)
            {
                MessageBox.Show("Error happened while updating 'comment' record(s).\nCould not save the changes.\n"+ exc.Message);
            }
        }

        /// <summary>
        /// if we want close the windows checks if there is any unsaved changes and ask for confirmation if there is, otherwise coses the window
        /// </summary>
        /// <param name="sender">button</param>
        /// <param name="e">eventargs</param>
        private void exitBtn_Click(object sender, EventArgs e)
        {
            if(this.excelButtonDataSet.MNBButtonLogs.GetChanges() != null)
            {
                var result = MessageBox.Show("You have unsaved changes. Are you sure you want to exit?", "Warning", MessageBoxButtons.YesNo);
                if (result != DialogResult.Yes)
                    return;
            }
            this.excelButtonDataSet.MNBButtonLogs.RejectChanges();
            this.Close();
        }
    }
}
