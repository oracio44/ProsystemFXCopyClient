using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string sWFX32;

        public Form1()
        {
            InitializeComponent();
            sWFX32 = GetWFX32();
        }

        //take command line arguments
        public Form1(string[] arg)
        {
            
            InitializeComponent();
            sWFX32 = GetWFX32();
            string fileName;
            foreach (string s in arg)
            {
                fileName = GetFileCode(s);
                addFile(fileName);
            }
        }

        //Add file name to listbox
        void addFile(string sName)
        {
            if (sName == "")
                return;
            else
            {
                if (!listBox1.Items.Contains(sName))
                {
                    listBox1.Items.Add(sName);
                }
            }
        }

        //Get WFX32 location.  Check NetINI location first, then Dir location
        string GetWFX32 ()
        {
            string strWFX32;
            strWFX32 = Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\CCHWinFx", "NetIniLocation", "none").ToString();
            if (strWFX32 == "none")
                strWFX32 = Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\CCHWinFx", "Dir", "none").ToString();
            strWFX32 = strWFX32.TrimEnd('\0');
            label4.Text = strWFX32;
            if (strWFX32 == "none")
            {
                MessageBox.Show("Please Run Workstation Setup before running this program");
                Application.Exit();
            }
            return strWFX32;
        }

        //Simplify file name to get 7 character working client code
        string GetFileCode(string strFullPath)
        {
            string strFileName = "";
            //verify file is of the expected file type for Prosystem .C00 files
            if (!strFullPath.Contains('.'))
            {
                MessageBox.Show(strFullPath + " is not a valid file");
                return strFileName;
            }
            else
            {
                if (!strFullPath.EndsWith("00"))
                {
                    MessageBox.Show(strFullPath + " is not a tax C00 file");
                    return strFileName;
                }
                else
                    strFileName = strFullPath.Substring(strFullPath.LastIndexOf('\\') + 1);
            }
            strFileName = strFileName.Substring(0, (strFileName.LastIndexOf('.') - 1));
            return strFileName;
        }

        bool SearchSub(string ClientCode)
        {
            string wfxNewBin = sWFX32 + "\\commun\\newbin\\";
            string wfxClient = sWFX32 + "\\client";
            string FileName = ClientCode + "*.U*";
            string sFilePath, sFileName;
            bool bFound = false;
            //Search for .U files using a provided client code in client directory
            foreach (string f in Directory.GetFiles(wfxClient, FileName, SearchOption.AllDirectories))
            {
                sFilePath = f;
                FileInfo fifile = new FileInfo(f);
                //check if found .u file is less than 12 KB
                if (fifile.Length < 12288)
                {
                    //check for presence of .B file, if none found, return to U file path
                    sFilePath = f.Replace(".U", ".B");
                    if (!File.Exists(sFilePath))
                    {
                        sFilePath = f;
                    }
                }
                sFileName = sFilePath.Substring(sFilePath.LastIndexOf('\\') + 1);
                listBox2.Items.Add(sFileName);
                sFileName = wfxNewBin + sFileName;
                File.Copy(sFilePath, sFileName, true);
                bFound = true;
            }
            //if Returns were found, return true.  If no returns found, return false
            return bFound;
        }

        //Accept drag-drop of files into first listbox
        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        //Actions to perform on files dropped into listbox
        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] dragFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
            string fileName;
            foreach (string s in dragFiles)
            {
                //Add files to listbox, avoiding duplicates
                if (!listBox1.Items.Contains(s))
                {
                    fileName = GetFileCode(s);
                    addFile(fileName);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sNone = "";
            progressBar1.Value = 0;
            progressBar1.Maximum = listBox1.Items.Count;
            progressBar1.Visible = true;
            //For each item in C00 List box, search for return files with that client code
            for (int i = 0; i<listBox1.Items.Count; i++)
            {
                progressBar1.Value++;
                if(!SearchSub(listBox1.Items[i].ToString()))
                    sNone = sNone + listBox1.Items[i] + "\n";
            }
            //Progressbar will be >0 if at least 1 file was attempted, report completed
            if (progressBar1.Maximum > 0)
                MessageBox.Show("Return files copied to " + sWFX32 + "\\commun\\newbin");
            //Show list of Clients with no returns found
            if (sNone.Length > 1)
            {
                sNone = sNone.Substring(0, sNone.Length - 1);
                MessageBox.Show("Unable to find return files for the following client(s):\n" + sNone);
            }
            progressBar1.Visible = false;
        }

        //Clear Listbox items
        private void button2_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
        }

        //Attempt to read Rebuild.log and take action on error files
        private void btRebuild_Click(object sender, EventArgs e)
        {
            string sRebuild, sLabel = "";
            string line, sClientFile;
            bool bLineE = false;
            List<string> sError = new List<string>();
            sRebuild = sWFX32 + "\\DATABASE\\Rebuild.log";
            try
            {
                listBox1.Items.Clear();
                listBox2.Items.Clear();
                FileStream file = new FileStream(sRebuild, FileMode.Open, FileAccess.Read);
                using (StreamReader Ereader = new StreamReader(file, Encoding.ASCII))
                {
                    //Set progress bar to expected max amount, and start at 0
                    progressBar1.Value = 0;
                    progressBar1.Maximum = File.ReadLines(sRebuild).Count();
                    progressBar1.Visible = true;
                    while ((line = Ereader.ReadLine()) != null)
                    {
                        // If flagged for second line is true, read a second line and append to current string
                        if (bLineE == true)
                        {
                            sLabel = sLabel + line.Substring(9);
                            sError.Add(sLabel);
                            bLineE = false;
                        }
                        //Start reading line, only Error lines will be reported, this should not happen if as a second line
                        if (line.Contains("ERROR: "))
                        {
                            sLabel = line;
                            //If line is max length, set flag for second line
                            if (sLabel.Length > 74)
                                bLineE = true;
                            sLabel = sLabel.Substring(9);
                            //if flagged for second line, remove "-" at end of line, and hold off adding to string list
                            if (bLineE)
                            {
                                if (sLabel[sLabel.Length - 1].Equals('-'))
                                    sLabel = sLabel.Substring(0, sLabel.Length - 1);
                            }
                            //if not flagged for a second line incoming, add line to string list
                            else
                                sError.Add(sLabel);
                        }
                        progressBar1.Value++;
                    }
                }
                file.Close();
                //Reset progress bar for next process, if no errors are returned, still setup progress bar for 1 maximum
                progressBar1.Value = 0;
                progressBar1.Maximum = sError.Count() + 1;
                sLabel = "";
                //Process each error line
                foreach (string s in sError)
                {
                    //Check if error is of expected format for this program to handle
                    if (s.Contains("Cannot revise to 2005"))
                    {
                        //Get file code, move the file to cannot process repository, search for return files, and add items to listbox
                        sClientFile = GetFileCode(s);
                        RemoveClient(sClientFile);
                        if (!listBox1.Items.Contains(sClientFile))
                            listBox1.Items.Add(sClientFile);
                        if (!SearchSub(sClientFile))
                            sLabel = sLabel + sClientFile + "\n";
                    }
                    progressBar1.Value++;
                }
                //Increment progress bar to completion, display appropriate message
                progressBar1.Value++;
                if (progressBar1.Value != 1)
                    MessageBox.Show("C00 files moved to " + sWFX32 + "\\client\\CannotRevise \n" + "Return Files copied to " + sWFX32 + "\\commun\\newbin");
                if (sLabel.Length > 1)
                {
                    sLabel = sLabel.Substring(0, sLabel.Length - 2);
                    MessageBox.Show("Unable to find return files for the following client(s):\n" + sLabel);
                }
                //Hide progress bar now that process is complete
                progressBar1.Visible = false;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //Move clients unable to revise to repository for storage and out of production
        void RemoveClient (string client)
        {
            string sCannotRevise = sWFX32 + "\\client\\CannotRevise";
            string sDestination = sCannotRevise + "\\" + client + "0.C00";
            string sRemove = sWFX32 + "\\client\\" + client + "0.C00";
            if (!Directory.Exists(sCannotRevise))
                Directory.CreateDirectory(sCannotRevise);
            if (!File.Exists(sRemove))
                return;
            if (File.Exists(sDestination))
                File.Delete(sDestination);
            File.Move(sRemove, sDestination);
        }

        private void Form1_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            MessageBox.Show("You can use this program to find client return files with the same name as a C00 file.\n\n" +
                "You can Drag-Drop C00 files into the \"C00 Files\" field.\n" +
                "Click \"Copy\" to search the WFX32\\Client folder for .U files.\n"+
                "Any .U files found will be copied to WFX32\\COMMUN\\newbin and listed in the \"Return Files\" list.\n\n" +
                "Clicking the \"Clear\" Button will clear the \"C00 Files\"and \"Return Files\"lists.\n\n" + 
                "\"Copy files from Rebuild.log\" will attempt to read Rebuild.Log and find any errors regarding \"Cannot Revise to 2005 Latest Format\"\n" + 
                "Found C00 files will be moved to WFX32\\Client\\CannotRevise, and Return files will be moved to WFX32\\COMMUN\\newbin");
        }
    }
}
