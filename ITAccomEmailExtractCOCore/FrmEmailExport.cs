using System;
using System.Data;
using ExcelDataReader;
using System.IO.Compression;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using Outlook2 = Microsoft.Office.Interop.Outlook;
using LINQhelper;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;
using System.Globalization;

namespace ITAccomEmailExtractC
{
    public partial class FrmEmailExport : Form
    {
        static string basePath = @"d:\outlook\emails\";
        static string servPath = @"d:\outlook\emails\";
        static string basePathProcess = @"d:\outlook\emails\";
        static string basePathTemp = @"d:\outlook\emails\";
        static string servPathProcess = @"d:\outlook\emails\";
        static int totalfilesize = 0;
        List<string> nList = new List<string>();
        List<string> nListRedo = new List<string>();


        public FrmEmailExport()
        {
            InitializeComponent();
        }



        private void btnGetMailBox_Click(object sender, EventArgs e)
        {
            // EnumerateAccounts();
            EnumerateFoldersInDefaultStore2();
        }

        private void EnumerateFoldersInDefaultStore()
            {
                Outlook2.Application Application = new Outlook2.Application();
                Outlook2.Folder root = Application.Session.DefaultStore.GetRootFolder() as Outlook2.Folder;
                EnumerateFolders(root);
            }


        private void EnumerateFoldersInDefaultStore2()
        {

            txtAttach.Text = "";
            txtAttach.Text += ("Outlook Attachment Extractor v0.1") + Environment.NewLine;
            txtAttach.Text += ("---------------------------------") + Environment.NewLine;
            int id;
        
            Outlook2.Application Application = new Outlook2.Application();
            Outlook2.Accounts accounts = Application.Session.Accounts;
            Outlook2.Stores newstores = Application.Session.Stores;
            Outlook2.Store currstore = null;



            id = 1;
           
            foreach (Outlook2.Account account in accounts)
            {
                txtAttach.Text += (id + ":" + EnumerateAccountEmailAddress(account)) + Environment.NewLine;
                //  Outlook2.Folder selectedFolder = Application.Session.DefaultStore.GetRootFolder() as Outlook2.Folder;
                //  selectedFolder = GetFolder(@"\\" + accounts[id].DisplayName);
                //  EnumerateFolders2(selectedFolder);
                lstBxAccount.Items.Add(EnumerateAccountEmailAddress(account));
                id++;
            }
            txtAttach.Text += ("Q: Quit Application") + Environment.NewLine;


            for (int i = 1; i <= newstores.Count; i++)

            {
                currstore = newstores[i];
                Outlook2.Folder root = currstore.GetRootFolder() as Outlook2.Folder;
                EnumerateFolders2(root);
            }
          //  Outlook2.Folder root =
           //     Application.Session.
            //   DefaultStore.GetRootFolder() as Outlook2.Folder;
          
           // EnumerateFolders2(root);
        }


        private void EnumerateFolders2(Outlook2.Folder folder)
        {
            Outlook2.Folders childFolders =
                folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook2.Folder childFolder in childFolders)
                {
                    // Write the folder path.
                    if (childFolder.FolderPath.Contains("Inbox"))
                    {
                        txtAttach.Text += (childFolder.FolderPath) + Environment.NewLine;
                        lstBxFolder.Items.Add(childFolder.FolderPath);
                        // string selectedItem = lstBxFolder.Items[lstBxFolder.SelectedIndex].ToString();
                        // Call EnumerateFolders using childFolder.
                        EnumerateFolders2(childFolder);
                    }


                }

            }
        }



        //create class to files from folder
        public class Files
        {
            public string FileName { get; set; }
            public string FilePath { get; set; }
        }
        //create list of files
        public List<Files> GetFiles(string path, string ftype)
        {
            List<Files> files = new List<Files>();
            foreach (string file in Directory.EnumerateFiles(path, "*." + ftype))
            {
                files.Add(new Files() { FileName = Path.GetFileName(file), FilePath = file });
            }
            return files;

        }





        private void GetandExtractZipFile()
        {
            string startPath;
            string zipPath;
            string extractPath;
            string extractFile;
            string[] extensionsArray = { ".zip", ".txt" };


            //create list of files
            List<Files> files = GetFiles(txtbxServerPath.Text + "TEMP\\", "zip");
            //create list of files names
            List<string> fileNames = files.Select(x => x.FileName).ToList();
            //create list of files paths
            List<string> filePaths = files.Select(x => x.FilePath).ToList();

            startPath = txtbxServerPath.Text + "TEMP\\";

            extractPath = txtbxServerPath.Text + "TEMP\\";

            foreach (string fileName in fileNames)
            {
                zipPath = txtbxServerPath.Text + "TEMP\\" + fileName;
                if (fileName.IndexOf(".zip", 0) > 0)
                {
                    //   MsgBoxResult result = Interaction.MsgBox("Do you want to extract " + extractPath + fileName.Replace(".zip", ".txt") + " ?", MsgBoxStyle.YesNo, "Extract Zip File");
                    extractFile = extractPath + fileName.Replace(".zip", ".txt");
                    if (!File.Exists(extractFile))
                    {
                        //MsgBoxResult result2 = Interaction.MsgBox("Do you want to extract " + txtbxServerPath.Text + "ABSATextFile\\" + fileName.Replace(".zip", ".txt") + " ?", MsgBoxStyle.YesNo, "Extract Zip File");

                        ZipFile.ExtractToDirectory(zipPath, extractPath);
                        // ZipFile.ExtractToDirectory(zipPath, basePathTemp);
                        txtAttach.Text += "Zip File Extracted Success " + fileName + Environment.NewLine;

                        if (!File.Exists(txtbxServerPath.Text + "Processed\\" + fileName))
                        {
                            File.Move(txtbxServerPath.Text + "TEMP\\" + fileName, txtbxServerPath.Text + "Processed\\" + fileName);
                        }
                        File.Delete(txtbxServerPath.Text + "TEMP\\" + fileName);

                    }
                    else

                    {
                        if (!File.Exists(txtbxServerPath.Text + "Processed\\" + fileName))
                        {
                            File.Move(txtbxServerPath.Text + "TEMP\\" + fileName, txtbxServerPath.Text + "Processed\\" + fileName);
                        }
                        File.Delete(txtbxServerPath.Text + "TEMP\\" + fileName);

                    }
                }


            }



            MessageBox.Show("Completed Adding Files to Unprocessed");



        }


        private void MoveFilestoABSA()
        {


            String[] fileSplit;
            String accFileName;
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //txtbxDesktoPath.Text = @"\\serverlaw5\Downloads\Balances\ABSA\";
            // basePath = path + @"\ABSA\DailyFiles\ALLFiles\";
            // basePathProcess = path + @"\ABSA\DailyFiles\ALLFiles\Processed\";
            // basePathTemp = path + @"\ABSA\DailyFiles\ALLFiles\Temp\";
            // txtbxServerPath.Text = @"\\serverlaw6b\LegABSA\ABSA\DialyFiles\ALLFiles\";
            // servPath = @"\\serverlaw6b\LegABSA\ABSA\DialyFiles\ALLFiles\";
            //  servPathProcess = @"\\serverlaw6b\LegABSA\ABSA\DialyFiles\ALLFiles\Processed\";

            if (!Directory.Exists(basePath))
            {
                Directory.CreateDirectory(basePath);
            }

            if (!Directory.Exists(servPath))
            {
                Directory.CreateDirectory(servPath);
            }
            if (!Directory.Exists(servPath + "TEMP\\"))
            {
                Directory.CreateDirectory(servPath + "TEMP\\");
            }

            if (!Directory.Exists(basePathProcess))
            {
                Directory.CreateDirectory(basePathProcess);
            }

            if (!Directory.Exists(servPathProcess))
            {
                Directory.CreateDirectory(servPathProcess);
            }

            //create list of files
            List<Files> files = GetFiles(txtbxDesktoPath.Text, "zip");
            //create list of files names
            List<string> fileNames = files.Select(x => x.FileName).ToList();
            //create list of files paths
            List<string> filePaths = files.Select(x => x.FilePath).ToList();

            
            
         
            foreach (string fileName in fileNames)
            {
                // check file does not exist in server
                if (!File.Exists(txtbxServerPath.Text + "TEMP\\" + fileName))
                {
                    File.Copy(txtbxDesktoPath.Text + fileName, txtbxServerPath.Text + "TEMP\\" + fileName);
                    txtAttach.Text += "Saving " + fileName + Environment.NewLine;
                }
                if (!File.Exists(txtbxDesktoPath.Text + "Done\\" + fileName))
                {
                    File.Move(txtbxDesktoPath.Text + fileName, txtbxDesktoPath.Text + "Done\\" + fileName);
                }
                File.Delete(txtbxDesktoPath.Text + fileName);



               
            }

        }
    



    private void btnEmailAtt_Click(object sender, EventArgs e)
        {

            {
                // EnumerateAccountsData();
                // EnumerateFoldersInDefaultStore3();
                //EnumerateFoldersInDefaultStore();
                //txtAttach.Text += ("Total file size:" + totalfilesize);
                // AddMewFileToList();


                MoveFilestoABSA();
                GetandExtractZipFile();
                AddNewFileToList();
                MessageBox.Show("ABSA SFTP Imported" );

            }
        }

        private void AddNewFileToList()
        {
            string[] fileSplit;

            DirectoryInfo d = new DirectoryInfo(txtbxServerPath.Text + "TEMP\\");
            FileInfo[] Files = d.GetFiles("*.txt"); //Getting Text files
            string str = "";
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            foreach (FileInfo filel in Files)
            {
                str = str + ", " + filel.Name;
                fileSplit = filel.Name.Split('_');

               
                var q = from p in BdSuiteContext.AbsaFileLogs
                        where p.Filename == filel.Name.Replace(".zip", ".txt")
                        select p;
                if (q.Count() == 0)
                {
                    if (fileSplit[0] != "ERROR")
                    {
                        if (fileSplit[1] == "OUTPUT")
                        {
                            AbsaFileLog objFilelog = new AbsaFileLog();
                            if (fileSplit[3] == "RECALL" || fileSplit[3] == "PAYMENT" || fileSplit[3] == "ACCOUNT")
                            {

                                objFilelog.Filename = filel.Name;
                                objFilelog.AccType = fileSplit[3];
                                objFilelog.FileDate = fileSplit[4];
                                objFilelog.Book = fileSplit[2];
                                objFilelog.CompPath = basePath;
                                objFilelog.ServPath = servPath;
                                objFilelog.Processed = 0;
                            }
                            else
                            {

                                objFilelog.Filename = filel.Name;
                                objFilelog.AccType = "ACCOUNT";
                                objFilelog.FileDate = fileSplit[3];
                                objFilelog.Book = fileSplit[2];
                                objFilelog.CompPath = basePath;
                                objFilelog.ServPath = servPath;
                                objFilelog.Processed = 0;

                            }
                            //BdSuiteContext.SubmitChanges();
                            try
                            {
                                BdSuiteContext.AbsaFileLogs.InsertOnSubmit(objFilelog);
                                BdSuiteContext.SubmitChanges();
                                 if (!File.Exists(txtbxServerPath.Text + "ABSATextFile\\" + filel.Name))
                                {
                                    File.Move(txtbxServerPath.Text + "TEMP\\" + filel.Name, txtbxServerPath.Text + "ABSATextFile\\" + filel.Name);
                                }
                                File.Delete(txtbxServerPath.Text + "TEMP\\" + filel.Name);
                                txtAttach.Text += "Filename Added  to Database" + Environment.NewLine;
                            }
                            catch (System.Data.Linq.DuplicateKeyException ex)
                            {
                                txtAttach.Text +=  ex.Message;
                                if (!File.Exists(txtbxServerPath.Text + "ABSATextFile\\" + filel.Name))
                                {
                                    File.Move(txtbxServerPath.Text + "TEMP\\" + filel.Name, txtbxServerPath.Text + "ABSATextFile\\" + filel.Name);
                                }
                                File.Delete(txtbxServerPath.Text + "TEMP\\" + filel.Name);
                                txtAttach.Text +=  "Filename  Already Exist" + Environment.NewLine;
                            }
                        }
                    }
                }
                else
                {
                    if (!File.Exists(txtbxServerPath.Text + "ABSATextFile\\" + filel.Name))
                    {
                        File.Move(txtbxServerPath.Text + "TEMP\\" + filel.Name, txtbxServerPath.Text + "ABSATextFile\\" + filel.Name);
                    }
                    File.Delete(txtbxServerPath.Text + "TEMP\\" + filel.Name);
                    txtAttach.Text += "Filename  Already Exist" + Environment.NewLine;
                }
            }

        }




        private void EnumerateFoldersInDefaultStore3()
        {

      
            Outlook2.Application Application = new Outlook2.Application();
            Outlook2.Accounts accounts = Application.Session.Accounts;
            Outlook2.Stores newstores = Application.Session.Stores;
            Outlook2.Store currstore = null;



            for (int i = 1; i <= newstores.Count; i++)

            {
                currstore = newstores[i];
                Outlook2.Folder root = currstore.GetRootFolder() as Outlook2.Folder;
                EnumerateFolders3(root);
            }
        }


        // Uses recursion to enumerate Outlook subfolders.
        private void EnumerateFolders3(Outlook2.Folder folder)
        {
            Outlook2.Folders childFolders =
                folder.Folders;
            

          

            if (childFolders.Count > 0)
            {
                foreach (Outlook2.Folder childFolder in childFolders)
                {
                    // Write the folder path.
                    if (childFolder.FolderPath.Contains("Inbox"))
                    {
                        txtAttach.Text += (childFolder.FolderPath) + Environment.NewLine;
                        lstBxFolder.Items.Add(childFolder.FolderPath);
                        //string selectedItem = lstBxFolder.Items[lstBxFolder.SelectedIndex].ToString();
                        // Call EnumerateFolders using childFolder.
                        if (childFolder.FolderPath.Contains(txtbxResponse.Text))
                        {
                            IterateMessages(childFolder,childFolder);
                            break;
                        }
                    }
                    EnumerateFolders3(childFolder);
                    
                }
               
            }
        }


       

        // Uses recursion to enumerate Outlook subfolders.
        private void EnumerateFolders(Outlook2.Folder folder)
            {
                Outlook2.Folders childFolders = folder.Folders;
                if (childFolders.Count > 0)
                {
                    foreach (Outlook2.Folder childFolder in childFolders)
                    {
                        // We only want Inbox folders - ignore Contacts and others
                        if (childFolder.FolderPath.Contains("Inbox"))
                        {
                        // Write the folder path.
                        txtAttach.Text +=  (childFolder.FolderPath);
                            // Call EnumerateFolders using childFolder, to see if there are any sub-folders within this one
                            EnumerateFolders(childFolder);
                        }
                    }
                }
                txtAttach.Text +=  ("Checking in " + folder.FolderPath);
                IterateMessages(folder,folder);
            }

            private void IterateMessages(Outlook2.Folder folder, Outlook2.Folder folderm)
            {
                // attachment extensions to save
//                string[] extensionsArray = { ".pdf", ".doc", ".xls", ".ppt", ".vsd", ".zip", ".rar", ".txt", ".csv", ".proj",".txt",".xlsx" };
            string[] extensionsArray = { ".zip",  ".txt" };

            // Iterate through all items ("messages") in a folder
            int iemailc = 0;
            int iattach = 0;
            string sTime;
            string sfrom;
            string sflname;
            string startPath;
            string zipPath;
            string extractPath;
            string[] fileSplit;

       
            txtbxAttach.Text = "0";
            txtbxMail.Text = "0";
              var fi = folder.Items.Restrict("[Unread]=true");
            if (fi != null)
                {

                  try
                  {
                    iemailc = 1;
                    foreach (Object item in fi)
                    {
                        iemailc++;
                        txtbxMail.Text = iemailc.ToString();
                        Outlook2.MailItem mi = (Outlook2.MailItem)item;
                        sTime = "";
                        sfrom = "";
                      //  if (mi.UnRead == true)
                      //  {
                            sTime = mi.CreationTime.ToString();
                            sfrom = mi.SenderEmailAddress.ToString();
                            sflname = sTime + "-" +  sfrom;
                            sflname = sflname.Replace(@"/", "").Replace(":", "").Replace(".", "_");
                        txtAttach.Text += " Busy with " + iemailc.ToString() + Environment.NewLine;
                        // if (!Directory.Exists(basePath + folder.FolderPath + @"\msg\"))
                        // {
                        //     Directory.CreateDirectory(basePath + folder.FolderPath + @"\msg\");
                        // }
                        // if (!File.Exists(basePath + folder.FolderPath + @"\msg\" + sflname + ".msg"))
                        // {
                        //     mi.SaveAs(basePath + folder.FolderPath + @"\msg\"  + sflname + ".msg", Outlook2.OlSaveAsType.olMSG);
                        // }

                        if (!Directory.Exists(basePath + @"\msg\"))
                             {
                                 Directory.CreateDirectory(basePath + @"\msg\");
                             }
                             if (!File.Exists(basePath + @"\msg\" + sflname + ".msg"))
                             {
                         // remove to save Desktop email
                        //         mi.SaveAs(basePath + @"\msg\"  + sflname + ".msg", Outlook2.OlSaveAsType.olMSG);
                             }

                            if (!Directory.Exists(servPath + @"\msg\"))
                            {
                                Directory.CreateDirectory(servPath + @"\msg\");
                            }
                            if (!File.Exists(servPath + @"\msg\" + sflname + ".msg"))

                            {
                        // remove to save Server email
                            //    mi.SaveAs(servPath + @"\msg\" + sflname + ".msg", Outlook2.OlSaveAsType.olMSG);
                            }




                            var attachments = mi.Attachments;
                            if (attachments.Count != 0)
                             {
                                iattach++;
                                txtbxAttach.Text = iattach.ToString();
                                // Create a directory to store the attachment 
                                //        if (!Directory.Exists(basePath + folder.FolderPath))
                                //    {
                                //        Directory.CreateDirectory(basePath + folder.FolderPath);
                                //    }

                                    if (!Directory.Exists(basePath ))
                                {
                                    Directory.CreateDirectory(basePath);
                                }

                                if (!Directory.Exists(servPath ))
                                {
                                    Directory.CreateDirectory(servPath );
                                }
                                if (!Directory.Exists(basePathProcess))
                                {
                                    Directory.CreateDirectory(basePathProcess);
                                }

                                if (!Directory.Exists(servPathProcess))
                                {
                                    Directory.CreateDirectory(servPathProcess);
                                }
                                //txtAttach.Text += (mi.Sender.Address);
                                //txtAttach.Text += (mi.Subject + " [" + attachments.Count + "]");
                                //txtAttach.Text += (generateFolder(folder.FolderPath, mi.Sender.Address));
                                for (int i = 1; i <= mi.Attachments.Count; i++)
                                {
                                    var fn = mi.Attachments[i].FileName.ToLower().Substring(mi.Attachments[i].FileName.Length-5,5);
                                    //check wither any of the strings in the extensionsArray are contained within the filename
                                    if (extensionsArray.Any(fn.Contains))
                                    {

                                        // Create a further sub-folder for the sender
                                        //-  if (!Directory.Exists(basePath + folder.FolderPath))
                                        //-  {
                                        //-      Directory.CreateDirectory(basePath + folder.FolderPath);
                                        //-  }
                                        totalfilesize +=  mi.Attachments[i].Size;
                                        //  if (!File.Exists(basePath + folder.FolderPath + @"\" + mi.Attachments[i].FileName))
                                        //   {
                                        //      txtAttach.Text += "Saving " + mi.Attachments[i].FileName + Environment.NewLine;
                                        //      mi.Attachments[i].SaveAsFile(basePath + folder.FolderPath + @"\" + mi.Attachments[i].FileName);
                                        //mi.Attachments[i].Delete();
                                        //  }
                                        //  else
                                        //  {
                                        txtAttach.Text += ("Processing ... " + mi.Attachments[i].FileName) + Environment.NewLine;
                                        //  }
                                        lstBxAccount.Items.Add(mi.Attachments[i].FileName.Replace(".zip", ".txt"));

                                        if (!File.Exists(basePathProcess + mi.Attachments[i].FileName))
                                        {
                                            
                                            fileSplit = mi.Attachments[i].FileName.Split('_');
                                           

                                            if (!File.Exists(basePath + mi.Attachments[i].FileName))
                                            {
                                                txtAttach.Text += "Saving " + mi.Attachments[i].FileName + Environment.NewLine;
                                               mi.Attachments[i].SaveAsFile(basePath + mi.Attachments[i].FileName);

                                            if (mi.Attachments[i].FileName.IndexOf(".txt", 0) > 0)
                                            {
                                                mi.Attachments[i].SaveAsFile(basePathTemp + mi.Attachments[i].FileName);
                                            }
                                           

                                            if (mi.Attachments[i].FileName.IndexOf(".zip", 0) > 0)
                                                {

                                                    if (!File.Exists(basePathProcess + mi.Attachments[i].FileName))
                                                    {
                                                        mi.Attachments[i].SaveAsFile(basePathProcess + mi.Attachments[i].FileName);
                                                       


                                                        startPath = basePath;
                                                        zipPath = basePath + mi.Attachments[i].FileName;
                                                        extractPath = basePath;
                                                        //      ZipFile.CreateFromDirectory(startPath, zipPath);
                                                        if (!File.Exists(basePath + mi.Attachments[i].FileName.Replace(".zip", ".txt")))
                                                        {
                                                            ZipFile.ExtractToDirectory(zipPath, extractPath);
                                                        ZipFile.ExtractToDirectory(zipPath, basePathTemp);
                                                        txtAttach.Text += "Zip File Extracted Success " + mi.Attachments[i].FileName + Environment.NewLine;
                                                        }
                                                    }
                                                }
                                                //mi.Attachments[i].Delete();
                                            }
                                            else
                                            {
                                                txtAttach.Text += ("Already saved " + mi.Attachments[i].FileName) + Environment.NewLine;
                                            }
                                        }
                                        else
                                        {
                                            txtAttach.Text += ("Already saved " + mi.Attachments[i].FileName) + Environment.NewLine;
                                        }
                                        

                                        if (!File.Exists(servPathProcess + mi.Attachments[i].FileName))
                                        {
                                            if (!File.Exists(servPath + mi.Attachments[i].FileName))
                                            {
                                                txtAttach.Text += "Saving " + mi.Attachments[i].FileName + Environment.NewLine;
                                            mi.Attachments[i].SaveAsFile(servPath + mi.Attachments[i].FileName);
                                            if (mi.Attachments[i].FileName.IndexOf(".txt", 0) > 0)
                                            {
                                               //  mi.Attachments[i].SaveAsFile(servPath + mi.Attachments[i].FileName);
                                               // mi.Attachments[i].SaveAsFile(basePathTemp + mi.Attachments[i].FileName);
                                            }
                                               
                                                //mi.Attachments[i].Delete();
                                                if (mi.Attachments[i].FileName.IndexOf(".zip", 0) > 0)
                                                {
                                                    if (!File.Exists(servPathProcess + mi.Attachments[i].FileName))
                                                    {
                                                        mi.Attachments[i].SaveAsFile(servPathProcess + mi.Attachments[i].FileName);
                                                        startPath = servPath;
                                                        zipPath = servPath + mi.Attachments[i].FileName;
                                                        extractPath = servPath;
                                                        //      ZipFile.CreateFromDirectory(startPath, zipPath);
                                                        if (!File.Exists(servPath + mi.Attachments[i].FileName.Replace(".zip",".txt")))
                                                        {
                                                            ZipFile.ExtractToDirectory(zipPath, extractPath);
                                                       // ZipFile.ExtractToDirectory(zipPath, basePathTemp);
                                                        txtAttach.Text += "Zip File Extracted Success " + mi.Attachments[i].FileName + Environment.NewLine;
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                txtAttach.Text += ("Already saved " + mi.Attachments[i].FileName) + Environment.NewLine;
                                            }
                                        }
                                        else
                                        {
                                            txtAttach.Text += ("Already saved " + mi.Attachments[i].FileName) + Environment.NewLine;
                                        }
                                    }
                                }
                           }

                       //     mi.UnRead = false;
                        //    mi.Save();
                        //    mi.Move(folderm);

                     //  }
                    }
                    iemailc = 1;
                    for (int i = fi.Count; i >= 1; i--)
                    {
                        iemailc++;
                        var mail = fi[i];

                        if (mail.UnRead)
                        {
                            mail.UnRead = false;
                            mail.Save();
                        }
                        txtAttach.Text += " Marked Unread -  " + iemailc.ToString() + Environment.NewLine;

                    }



                }
                   catch (Exception e)
                  {
                        txtAttach.Text += "An error occurred: '{0}'" + e.ToString() + Environment.NewLine;
                  }
                }
            }

            // Retrieves the email address for a given account object
            private string EnumerateAccountEmailAddress(Outlook2.Account account)
            {
                try
                {
                    if (string.IsNullOrEmpty(account.SmtpAddress) || string.IsNullOrEmpty(account.UserName))
                    {
                        Outlook2.AddressEntry oAE = account.CurrentUser.AddressEntry as Outlook2.AddressEntry;
                        if (oAE.Type == "EX")
                        {
                            Outlook2.ExchangeUser oEU = oAE.GetExchangeUser() as Outlook2.ExchangeUser;
                            return oEU.PrimarySmtpAddress;
                        }
                        else
                        {
                            return oAE.Address;
                        }
                    }
                    else
                    {
                        return account.SmtpAddress;
                    }
                }
                catch (Exception ex)
                {
                    txtAttach.Text += (ex.Message);
                    return "";
                }
            }

        private void EnumerateAccounts()
        {
            txtAttach.Text = "";
            txtAttach.Text += ("Outlook Attachment Extractor v0.1") + Environment.NewLine;
            txtAttach.Text += ("---------------------------------") + Environment.NewLine;
            int id;
            Outlook2.Application Application = new Outlook2.Application();
            Outlook2.Accounts accounts = Application.Session.Accounts;

          


            id = 1;
            foreach (Outlook2.Account account in accounts)
            {
                txtAttach.Text += (id + ":" + EnumerateAccountEmailAddress(account)) + Environment.NewLine;
                lstBxAccount.Items.Add(EnumerateAccountEmailAddress(account));
                id++;
            }
            txtAttach.Text +=  ("Q: Quit Application");
        }
            private void EnumerateAccountsData()
            {
            int id;
            Outlook2.Application Application = new Outlook2.Application();
            Outlook2.Accounts accounts = Application.Session.Accounts;

            id = 1;

            foreach (Outlook2.Account account in accounts)
            {
                id++;
            }

            string response = "";
            response = txtbxResponse.Text.ToUpper();
                    if (response == "Q")
                    {
                        txtAttach.Text += ("Quitting");
                        return;
                    }
                    if (response != "")
                    {
                        if (Int32.Parse(response.Trim()) >= 1 && Int32.Parse(response.Trim()) < id)
                        {
                            txtAttach.Text += ("Processing: " + accounts[Int32.Parse(response.Trim())].DisplayName);
                            txtAttach.Text += ("Processing: " + EnumerateAccountEmailAddress(accounts[Int32.Parse(response.Trim())]));

                            Outlook2.Folder selectedFolder = Application.Session.DefaultStore.GetRootFolder() as Outlook2.Folder;
                            selectedFolder = GetFolder(@"\\" + accounts[Int32.Parse(response.Trim())].DisplayName);
                            EnumerateFolders(selectedFolder);
                            txtAttach.Text += ("Finished Processing " + accounts[Int32.Parse(response.Trim())].DisplayName);
                            txtAttach.Text += ("");
                        }
                        else
                        {
                            txtAttach.Text += ("Invalid Account Selected");
                        }
                    }
                

            }

            // Returns Folder object based on folder path
            private Outlook2.Folder GetFolder(string folderPath)
            {
                txtAttach.Text += ("Looking for: " + folderPath);
                Outlook2.Folder folder;
                string backslash = @"\";
                try
                {
                    if (folderPath.StartsWith(@"\\"))
                    {
                        folderPath = folderPath.Remove(0, 2);
                    }
                    String[] folders = folderPath.Split(backslash.ToCharArray());
                    Outlook2.Application Application = new Outlook2.Application();
                    folder = Application.Session.Folders[folders[0]] as Outlook2.Folder;
                    if (folder != null)
                    {
                        for (int i = 1; i <= folders.GetUpperBound(0); i++)
                        {
                            Outlook2.Folders subFolders = folder.Folders;
                            folder = subFolders[folders[i]] as Outlook2.Folder;
                            if (folder == null)
                            {
                                return null;
                            }
                        }
                    }
                    return folder;
                }
                catch (Exception ex)
                {
                    txtAttach.Text += (ex.Message);
                    return null;
                }
            }

        private void txtbxResponse_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void FrmEmailExport_Load(object sender, EventArgs e)
        {
           
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            txtbxDesktoPath.Text = @"\\serverlaw5\Downloads\Balances\ABSA\";
            basePath = path + @"\ABSA\DailyFiles\ALLFiles\";
            basePathProcess = path + @"\ABSA\DailyFiles\ALLFiles\Processed\";
            basePathTemp = path + @"\ABSA\DailyFiles\ALLFiles\Temp\";
            txtbxServerPath.Text = @"\\serverlaw6b\LegABSA\ABSA\DialyFiles\ALLFiles\";
            servPath = @"\\serverlaw6b\LegABSA\ABSA\DialyFiles\ALLFiles\";
            servPathProcess = @"\\serverlaw6b\LegABSA\ABSA\DialyFiles\ALLFiles\Processed\";

        }

            private void FrmEmailExport_FormClosed(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnFileProcess_Click(object sender, EventArgs e)
        {
            string fileName = txtlstBxAccount.Text;
                //lstBxAccount.SelectedItem.ToString();    //.GetItemText(lstBxAccount.SelectedValue);
           ProcessFiles(fileName,basePath);

        }

        protected void ProcessFiles(String processFileName ,String NewPath)

        {
            String[] fileSplit;
           
            fileSplit = processFileName.Split('_');
            if (fileSplit.Length > 2)
            {
                if (fileSplit[3] == "ACCOUNT")
                {
                    ProcessAccFile(processFileName, NewPath);
                }
                if (fileSplit[3] == "RECALL")
                {
                    ProcessRecallFile(processFileName, NewPath);
                }
                if (fileSplit[3] == "PAYMENT")
                {
                    ProcessPayFile(processFileName, NewPath);
                }
                if (fileSplit[3] != "ACCOUNT"  || fileSplit[3] == "RECALL" || fileSplit[3] == "PAYMENT")
                {
                    ProcessAccFile2(processFileName, NewPath);
                }
            }
        }


            protected void ProcessAccFile(String accFileName,String NewPath)
        {
            String[] fileSplit;
            int noFields;
            int noUpd = 0;
            int allCount = 0;
            String newdata;
          
            DateTime dateNew;

            fileSplit = accFileName.Split('_');

            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();


            ///string accdelStatement = "Delete from ABSAAccFile where Filename = " + accFileName.Replace(".zip", ".txt");
            ///BdSuiteContext.ExecuteQuery<ABSAAccFile>(accdelStatement);
            ///
            string accdelStatement = "Delete from ABSAAccFile where Filename = {0},\"" + accFileName.Replace(".zip", ".txt") + "\"";
           /// BdSuiteContext.ExecuteCommand(accdelStatement);      
           /// BdSuiteContext.ExecuteQuery<ABSAAccFile>(accdelStatement);


            int recordsToInsert = int.Parse(tbRecordsToInsert.Text);
            this.Cursor = Cursors.WaitCursor;

             List<ABSAAccFile> recordsToBeInserted = new List<ABSAAccFile>();
            using (DataClassBDSuiteDataContext dc = new DataClassBDSuiteDataContext())
            {
                using (var textFieldParser = new TextFieldParser(NewPath + accFileName))
                {
                    textFieldParser.TextFieldType = FieldType.Delimited;
                    textFieldParser.Delimiters = new[] { "|" };
               
                    textFieldParser.TrimWhiteSpace = true;
                    textFieldParser.HasFieldsEnclosedInQuotes = false;
                    // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
                    while (!textFieldParser.EndOfData)
                    {
                        try
                        {
                            noUpd++;
                            allCount++;
                            string[] parts = textFieldParser.ReadFields();
                            noFields = parts.Length;
                            txtProcess.Text = accFileName + " - " + noFields.ToString();
                            ABSAAccFile newAccfile = new ABSAAccFile();

                            newAccfile.Filename = accFileName.Replace(".zip", ".txt");
                            newAccfile.AccType = fileSplit[3];
                            newAccfile.FileDate = fileSplit[4];
                            dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        //    newAccfile.DateF = Convert.tol(dateNew) + 2;
                            newAccfile.Book = fileSplit[2];
                            newAccfile.Ind = parts[0];
                            newAccfile.Accountnumber = parts[1];
                            newAccfile.Title = parts[2];
                            newAccfile.Initials = parts[3];
                            newAccfile.Surname = parts[4];
                            newAccfile.IDNumber = parts[5];
                            newAccfile.EmailAddress = parts[6];
                            newAccfile.CellNo = parts[7];
                            newAccfile.HomeTel = parts[8];
                            newAccfile.WorkTel = parts[9];
                            newAccfile.FaxNo = parts[10];
                            newAccfile.HomeLine1 = parts[11];
                            newAccfile.HomeLine2 = parts[12];
                            newAccfile.HomeLine3 = parts[13];
                            newAccfile.HomeLine4 = parts[14];
                            newAccfile.HomePostal = parts[15];
                            newAccfile.PostalLine1 = parts[16];
                            newAccfile.PostalLine2 = parts[17];
                            newAccfile.PostalLine3 = parts[18];
                            newAccfile.PostalLine4 = parts[19];
                            newAccfile.PostalCode = parts[20];
                            newAccfile.EmployerName = parts[21];
                            newAccfile.EmployerAddLine1 = parts[22];
                            newAccfile.EmployerAddLine2 = parts[23];
                            newAccfile.EmployerAddLine3 = parts[24];
                            newAccfile.EmployerPostalCode = parts[25];
                            newAccfile.Occupation = parts[26];
                            newAccfile.Product = parts[27];
                            newAccfile.InstructionReceivedDate = parts[28];
                            newAccfile.InterestFromDate = parts[29];
                            newdata = parts[30];
                            if (!string.IsNullOrEmpty(parts[30]) || parts[30] == "0.00")
                            {
                                newAccfile.InterestRate = Convert.ToDecimal(parts[30]);
                            }
                            newAccfile.SBUCode = parts[31];
                            newAccfile.BranchNr = parts[32];
                            newAccfile.LastPaymentDate = parts[33];
                            if (!string.IsNullOrEmpty(parts[34]) || parts[34] == "0.00")
                            {

                                newAccfile.LastPaymentAmount = Convert.ToDecimal(parts[34]);
                            }
                            newAccfile.Tag1 = parts[35];
                            newAccfile.Tagassignmentdate1 = parts[36];
                            newAccfile.Tag2 = parts[37];
                            newAccfile.Tagassignmentdate2 = parts[38];
                            newAccfile.Tag3 = parts[39];
                            newAccfile.Tagassignmentdate3 = parts[40];
                            newAccfile.Tag4 = parts[41];
                            newAccfile.Tagassignmentdate4 = parts[42];
                            newAccfile.Tag5 = parts[43];
                            newAccfile.Tagassignmentdate5 = parts[44];
                            newAccfile.SummonsCourt = parts[45];
                            newAccfile.SummonsCaseNumber = parts[46];
                            newAccfile.SummonsServedDate = parts[47];
                            newAccfile.JudgementCaseNumber = parts[48];
                            newAccfile.JudgementCourtDate = parts[49];
                            newAccfile.LegalJudgementDAte = parts[50];
                            if (!string.IsNullOrEmpty(parts[51]))
                            {
                                newAccfile.DefaultJudgementAmount = Convert.ToDecimal(parts[51]);
                            }
                            newAccfile.SIEOUTCOMEDateSold = parts[52];
                            if (!string.IsNullOrEmpty(parts[53]))
                            {

                                newAccfile.SIEOUTCOMESoldAmount = Convert.ToDecimal(parts[53]);
                            }
                            newAccfile.DateWrittenOff = parts[54];
                            if (!string.IsNullOrEmpty(parts[55]))
                            {
                                newAccfile.CapitalWrittenOff = Convert.ToDecimal(parts[55]);
                            }
                            if (!string.IsNullOrEmpty(parts[56]))
                            {
                                newAccfile.InterestWrittenOff = Convert.ToDecimal(parts[56]);
                            }
                            newAccfile.RecoverySummonsCourt = parts[57];
                            newAccfile.RecoverySummonsCase = parts[58];
                            newAccfile.RecoverySummonsServedDate = parts[59];
                            newAccfile.RecoveryJudgementDate = parts[60];
                            newAccfile.RecoverySIEOUTCOMEDateSold = parts[61];
                            if (!string.IsNullOrEmpty(parts[62]))
                            {
                                newAccfile.RecoverySIEOUTCOMESoldAmount = Convert.ToDecimal(parts[62]);
                            }
                            if (!string.IsNullOrEmpty(parts[63]))
                            {
                                newAccfile.NettCredits = Convert.ToDecimal(parts[63]);
                            }
                            if (!string.IsNullOrEmpty(parts[64]))
                            {

                                newAccfile.Arrears = Convert.ToDecimal(parts[64]);
                            }



                            if (!string.IsNullOrEmpty(parts[65]))
                            {
                                newAccfile.Balance = Convert.ToDecimal(parts[65]);
                            }
                            newAccfile.Asset = parts[66];
                            if (noFields > 67)
                            {
                                newAccfile.PrelegalInstructionReceivedDate = parts[67];
                                newAccfile.PrelegalName = parts[68];
                                newAccfile.LegalInstructionReceivedDate = parts[69];
                                newAccfile.LegalName = parts[70];
                            }
                            if (noFields > 71)
                            {


                                if (!string.IsNullOrEmpty(parts[71]))
                                {
                                    newAccfile.InstalmentAmount = Convert.ToDecimal(parts[71]);
                                }
                                newAccfile.InstalmentDueDate = parts[72];
                                newAccfile.InstalmentDate = parts[73];
                                newAccfile.DateaccountOpened = parts[74];
                                newAccfile.NID = parts[75];
                                newAccfile.NDD = parts[76];
                                newAccfile.MOB = parts[77];
                                if (!string.IsNullOrEmpty(parts[78]))
                                {
                                    newAccfile.OriginalBalance = Convert.ToDecimal(parts[78]);
                                }
                                if (!string.IsNullOrEmpty(parts[79]))
                                {

                                    newAccfile.OutstandingBalance = Convert.ToDecimal(parts[79]);
                                }
                                newAccfile.Tag6 = parts[80];
                                newAccfile.Tag6AssignmentDate = parts[81];
                                newAccfile.Tag7 = parts[82];
                                newAccfile.Tag7AssignmentDate = parts[83];
                                newAccfile.Tag8 = parts[84];
                                newAccfile.Tag8AssignmentDate = parts[85];
                                if (!string.IsNullOrEmpty(parts[86]))
                                {
                                    newAccfile.IntRate = Convert.ToDecimal(parts[86]);
                                }
                                newAccfile.LoanTerm = parts[87];
                                newAccfile.NrPTPCCC = parts[88];
                                newAccfile.NrPTPConsec = parts[89];
                                newAccfile.NrPTPLife = parts[90];
                                newAccfile.BranchName = parts[91];
                            }
                            if (noFields > 92)
                            {

                                newAccfile.RiskType = parts[92];
                                newAccfile.TRIADPool = parts[93];
                                newAccfile.CreditLifeapplicable = parts[94];
                                newAccfile.PlasticNr = parts[95];
                                newAccfile.LegalAccountNBumber = parts[96];
                                newAccfile.Unemployed = parts[97];
                                newAccfile.Unemployedsince = parts[98];
                                newAccfile.UnemploymentAffidavitDate = parts[99];
                                newAccfile.Section129Letter = parts[100];
                                newAccfile.varchargeoffDate = parts[101];
                                if (!string.IsNullOrEmpty(parts[102]))
                                {
                                    newAccfile.ArrearBalanceatvarchargeoff = Convert.ToDecimal(parts[102]);
                                }
                                if (!string.IsNullOrEmpty(parts[103]))
                                {
                                    newAccfile.OutstandingBalanceatvarchargeoff = Convert.ToDecimal(parts[103]);
                                }
                                if (!string.IsNullOrEmpty(parts[104]))
                                {
                                    newAccfile.Interestatvarchargeoff = Convert.ToDecimal(parts[104]);
                                }
                                if (!string.IsNullOrEmpty(parts[105]))
                                {

                                    newAccfile.HandoverAmount = Convert.ToDecimal(parts[105]);
                                }
                                newAccfile.AssetDescription = parts[106];
                                newAccfile.AssetRegistration = parts[107];
                                newAccfile.Enginenumber = parts[108];
                                newAccfile.Chassisnumber = parts[109];
                                newAccfile.Vinnumber = parts[110];
                                newAccfile.MMCode = parts[111];
                                if ( parts[112].Length > 30 )
                                {
                                    newAccfile.DealerName = parts[112].Substring(0,30);
                                }
                                else
                                {
                                    newAccfile.DealerName = parts[112];
                                }
                               

                                newAccfile.DealerNumber = parts[113];
                                if (!string.IsNullOrEmpty(parts[114]))
                                {
                                    newAccfile.TradeValue = Convert.ToDecimal(parts[114]);
                                }
                                if (!string.IsNullOrEmpty(parts[115]))
                                {

                                    newAccfile.RetailValue = Convert.ToDecimal(parts[115]);
                                }
                                if (!string.IsNullOrEmpty(parts[116]))
                                {
                                    newAccfile.SoldAmount = Convert.ToDecimal(parts[116]);
                                }
                                newAccfile.TradeCentreNumber = parts[117];
                                newAccfile.VehicleStockNumber = parts[118];
                                newAccfile.Assetinpossession = parts[119];
                                newAccfile.WarrantIssued = parts[120];
                                newAccfile.Warrant_Sheriff = parts[121].Trim();
                                newAccfile.Sheriffappointed = parts[122].Trim();
                                newAccfile.DocumentationCompleted = parts[123].Trim();
                                newAccfile.DisputeStatus = parts[124];
                                newAccfile.DisputeStatusDate = parts[125];
                                newAccfile.Complaint = parts[126];
                                newAccfile.Complaintdate = parts[127];
                                newAccfile.DebtReviewIndicator = parts[128];
                                newAccfile.InternalPlacementbefore = parts[129];
                                newAccfile.InternalPlacementdate = parts[130];
                                newAccfile.Prescriptionaccount = parts[131];
                                newAccfile.LastPTP = parts[132];
                            }
                            if (noFields > 133)
                            {
                                newAccfile.FirstName = parts[133];
                                if (!string.IsNullOrEmpty(parts[134]))
                                {
                                    newAccfile.SettlmentPercentage = Convert.ToDecimal(parts[134]);
                                }
                                if (!string.IsNullOrEmpty(parts[135]))
                                {
                                    newAccfile.PastDue = Convert.ToDecimal(parts[135]);
                                }
                                if (!string.IsNullOrEmpty(parts[136]))
                                {
                                    newAccfile.Limit = Convert.ToDecimal(parts[136]);
                                }
                                if (!string.IsNullOrEmpty(parts[137]))
                                {
                                    newAccfile.OverLimit = Convert.ToDecimal(parts[137]);
                                }
                                if (!string.IsNullOrEmpty(parts[138]))
                                {
                                    newAccfile.VAtAmt = Convert.ToDecimal(parts[138]);
                                }
                            }
                            if (noFields > 139)
                            {

                                if (!string.IsNullOrEmpty(parts[139]))
                                {
                                    newAccfile.ECBalance = Convert.ToDecimal(parts[139]);
                                }
                                if (!string.IsNullOrEmpty(parts[140]))
                                {
                                    newAccfile.CONSBalance = Convert.ToDecimal(parts[140]);
                                }
                                if (!string.IsNullOrEmpty(parts[141]))
                                {
                                    newAccfile.PrePrescribeTag = parts[141];
                                }
                            }
                            if (noFields > 142)
                            {
                                if (!string.IsNullOrEmpty(parts[142]))
                                {
                                    newAccfile.Total_OutStanding_Bal = Convert.ToDecimal(parts[142]);
                                }
                                if (!string.IsNullOrEmpty(parts[143]))
                                {
                                    newAccfile.OverDarft_Interest_Rate = Convert.ToDecimal(parts[143]);
                                }
                                if (!string.IsNullOrEmpty(parts[144]))
                                {
                                    newAccfile.Prev_Install_Amount = Convert.ToDecimal(parts[144]);
                                }
                                if (!string.IsNullOrEmpty(parts[145]))
                                {
                                    newAccfile.Decline_Reason = parts[145];
                                }
                                if (!string.IsNullOrEmpty(parts[146]))
                                {
                                    newAccfile.Memo = parts[146];
                                }
                                if (!string.IsNullOrEmpty(parts[147]))
                                {
                                    newAccfile.XDS_ITC1 = parts[147];
                                }
                                if (!string.IsNullOrEmpty(parts[148]))
                                {
                                    newAccfile.XDS_ITC2 = parts[148];
                                }
                                if (!string.IsNullOrEmpty(parts[149]))
                                {
                                    newAccfile.EDC_GIVEN1 = parts[149];
                                }
                                if (!string.IsNullOrEmpty(parts[150]))
                                {
                                    newAccfile.EDC_GIVEN2 = parts[150];
                                }
                                if (!string.IsNullOrEmpty(parts[151]))
                                {
                                    newAccfile.INTERNAL_PHN1 = parts[151];
                                }
                                if (!string.IsNullOrEmpty(parts[152]))
                                {
                                    newAccfile.INTERNAL_PHN2 = parts[152];
                                }
                                if (!string.IsNullOrEmpty(parts[153]))
                                {
                                    newAccfile.INTERNAL_PHN3 = parts[153];
                                }
                                if (!string.IsNullOrEmpty(parts[154]))
                                {
                                    newAccfile.INTERNAL_PHN4 = parts[154];
                                }
                                if (!string.IsNullOrEmpty(parts[155]))
                                {
                                    newAccfile.INTERNAL_PHN5 = parts[155];
                                }
                                if (!string.IsNullOrEmpty(parts[156]))
                                {
                                    newAccfile.INTERNAL_PHN6 = parts[156];
                                }
                            }
                            if (noFields > 169)

                            {

                                if (!string.IsNullOrEmpty(parts[157]))
                                {
                                    newAccfile.Placement = parts[157];
                                }
                                if (!string.IsNullOrEmpty(parts[158]))
                                {
                                    newAccfile.Delinquency5 = parts[158];
                                }
                                if (!string.IsNullOrEmpty(parts[159]))
                                {
                                    newAccfile.Delinquency30 = parts[159];
                                }
                                if (!string.IsNullOrEmpty(parts[160]))
                                {
                                    newAccfile.Delinquency60 = parts[160];
                                }
                                if (!string.IsNullOrEmpty(parts[161]))
                                {
                                    newAccfile.Delinquency90 = parts[161];
                                }
                                if (!string.IsNullOrEmpty(parts[162]))
                                {
                                    newAccfile.Delinquency120 = parts[162];
                                }
                                if (!string.IsNullOrEmpty(parts[163]))
                                {
                                    newAccfile.Delinquency150 = parts[163];
                                }
                                if (!string.IsNullOrEmpty(parts[164]))
                                {
                                    newAccfile.Delinquency180 = parts[164];
                                }
                                if (!string.IsNullOrEmpty(parts[165]))
                                {
                                    newAccfile.Delinquency210 = parts[165];
                                }
                                if (!string.IsNullOrEmpty(parts[166]))
                                {
                                    newAccfile.Latest_Payment_Schedule_Date = parts[166];
                                }
                                if (!string.IsNullOrEmpty(parts[167]))
                                {
                                    newAccfile.OD_NONOD = parts[167];
                                }
                                if (!string.IsNullOrEmpty(parts[168]))
                                {
                                    newAccfile.OverDraftLimit = parts[168];
                                }
                                if (!string.IsNullOrEmpty(parts[169]))
                                {
                                    newAccfile.Placement_Date = parts[169];
                                }
                        

                            }
                                recordsToBeInserted.Add(newAccfile);
                        }
                        catch (System.FormatException ex3)
                        {
                            txtAttach.Text +=  ex3.Message;
                            txtAttach.Text +=  string.Format("Error record not imported format error ") + Environment.NewLine;
                        }
                        if (noUpd >= recordsToInsert)
                        {
                                                     //  Insert this List<> of records using a Bulk Insert
                            DataTableHelper.BulkCopyToDatabase(recordsToBeInserted, "ABSAAccFile", dc);
                                                     
                            txtAttach.Text += string.Format("Inserted  " + noUpd.ToString() + " records TO AccFiles ") + Environment.NewLine;
                            recordsToBeInserted.Clear();
                            noUpd = 0;

                        }

                    }
                }
            

            DateTime dtInsertStarted = DateTime.Now;

            //  Insert this List<> of records using a Bulk Insert
            DataTableHelper.BulkCopyToDatabase(recordsToBeInserted, "ABSAAccFile", dc);

            TimeSpan duration = DateTime.Now - dtInsertStarted;
            txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", recordsToInsert, duration.Seconds, duration.Milliseconds) + Environment.NewLine;
            }
            this.Cursor = Cursors.Arrow;
        }

        protected void ProcessAccFile2(String accFileName, String NewPath)
        {
            String[] fileSplit;
            int noFields;
            int noUpd = 0;
            int allCount = 0;
            String newdata;

            DateTime dateNew;

            fileSplit = accFileName.Split('_');

            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();


            ///string accdelStatement = "Delete from ABSAAccFile where Filename = " + accFileName.Replace(".zip", ".txt");
            ///BdSuiteContext.ExecuteQuery<ABSAAccFile>(accdelStatement);
            ///
            string accdelStatement = "Delete from ABSAAccFile where Filename = {0},\"" + accFileName.Replace(".zip", ".txt") + "\"";
            /// BdSuiteContext.ExecuteCommand(accdelStatement);      
            /// BdSuiteContext.ExecuteQuery<ABSAAccFile>(accdelStatement);


            int recordsToInsert = int.Parse(tbRecordsToInsert.Text);
            this.Cursor = Cursors.WaitCursor;

            List<ABSAAccFile> recordsToBeInserted = new List<ABSAAccFile>();
            using (DataClassBDSuiteDataContext dc = new DataClassBDSuiteDataContext())
            {
                using (var textFieldParser = new TextFieldParser(NewPath + accFileName))
                {
                    textFieldParser.TextFieldType = FieldType.Delimited;
                    textFieldParser.Delimiters = new[] { "|" };
                    textFieldParser.TrimWhiteSpace = true;
                    textFieldParser.HasFieldsEnclosedInQuotes = true;
                    // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
                    while (!textFieldParser.EndOfData)
                    {
                        try
                        {
                            noUpd++;
                            allCount++;
                            string[] parts = textFieldParser.ReadFields();
                            noFields = parts.Length;
                            txtProcess.Text = accFileName + " - " + noFields.ToString();
                            ABSAAccFile newAccfile = new ABSAAccFile();

                            newAccfile.Filename = accFileName.Replace(".zip", ".txt");
                            newAccfile.AccType = "ACCOUNT";
                            newAccfile.FileDate = fileSplit[3];
                            dateNew = DateTime.ParseExact(fileSplit[3].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                            //    newAccfile.DateF = Convert.tol(dateNew) + 2;
                            newAccfile.Book = fileSplit[2];
                            newAccfile.Ind = parts[0];
                            newAccfile.Accountnumber = parts[1];
                            newAccfile.Title = parts[2];
                            newAccfile.Initials = parts[3];
                            newAccfile.Surname = parts[4];
                            newAccfile.IDNumber = parts[5];
                            newAccfile.EmailAddress = parts[6];
                            newAccfile.CellNo = parts[7];
                            newAccfile.HomeTel = parts[8];
                            newAccfile.WorkTel = parts[9];
                            newAccfile.FaxNo = parts[10];
                            newAccfile.HomeLine1 = parts[11];
                            newAccfile.HomeLine2 = parts[12];
                            newAccfile.HomeLine3 = parts[13];
                            newAccfile.HomeLine4 = parts[14];
                            newAccfile.HomePostal = parts[15];
                            newAccfile.PostalLine1 = parts[16];
                            newAccfile.PostalLine2 = parts[17];
                            newAccfile.PostalLine3 = parts[18];
                            newAccfile.PostalLine4 = parts[19];
                            newAccfile.PostalCode = parts[20];
                            newAccfile.EmployerName = parts[21];
                            newAccfile.EmployerAddLine1 = parts[22];
                            newAccfile.EmployerAddLine2 = parts[23];
                            newAccfile.EmployerAddLine3 = parts[24];
                            newAccfile.EmployerPostalCode = parts[25];
                            newAccfile.Occupation = parts[26];
                            newAccfile.Product = parts[27];
                            newAccfile.InstructionReceivedDate = parts[28];
                            newAccfile.InterestFromDate = parts[29];
                            newdata = parts[30];
                            if (!string.IsNullOrEmpty(parts[30]) || parts[30] == "0.00")
                            {
                                newAccfile.InterestRate = Convert.ToDecimal(parts[30]);
                            }
                            newAccfile.SBUCode = parts[31];
                            newAccfile.BranchNr = parts[32];
                            newAccfile.LastPaymentDate = parts[33];
                            if (!string.IsNullOrEmpty(parts[34]) || parts[34] == "0.00")
                            {

                                newAccfile.LastPaymentAmount = Convert.ToDecimal(parts[34]);
                            }
                            newAccfile.Tag1 = parts[35];
                            newAccfile.Tagassignmentdate1 = parts[36];
                            newAccfile.Tag2 = parts[37];
                            newAccfile.Tagassignmentdate2 = parts[38];
                            newAccfile.Tag3 = parts[39];
                            newAccfile.Tagassignmentdate3 = parts[40];
                            newAccfile.Tag4 = parts[41];
                            newAccfile.Tagassignmentdate4 = parts[42];
                            newAccfile.Tag5 = parts[43];
                            newAccfile.Tagassignmentdate5 = parts[44];
                            newAccfile.SummonsCourt = parts[45];
                            newAccfile.SummonsCaseNumber = parts[46];
                            newAccfile.SummonsServedDate = parts[47];
                            newAccfile.JudgementCaseNumber = parts[48];
                            newAccfile.JudgementCourtDate = parts[49];
                            newAccfile.LegalJudgementDAte = parts[50];
                            if (!string.IsNullOrEmpty(parts[51]))
                            {
                                newAccfile.DefaultJudgementAmount = Convert.ToDecimal(parts[51]);
                            }
                            newAccfile.SIEOUTCOMEDateSold = parts[52];
                            if (!string.IsNullOrEmpty(parts[53]))
                            {

                                newAccfile.SIEOUTCOMESoldAmount = Convert.ToDecimal(parts[53]);
                            }
                            newAccfile.DateWrittenOff = parts[54];
                            if (!string.IsNullOrEmpty(parts[55]))
                            {
                                newAccfile.CapitalWrittenOff = Convert.ToDecimal(parts[55]);
                            }
                            if (!string.IsNullOrEmpty(parts[56]))
                            {
                                newAccfile.InterestWrittenOff = Convert.ToDecimal(parts[56]);
                            }
                            newAccfile.RecoverySummonsCourt = parts[57];
                            newAccfile.RecoverySummonsCase = parts[58];
                            newAccfile.RecoverySummonsServedDate = parts[59];
                            newAccfile.RecoveryJudgementDate = parts[60];
                            newAccfile.RecoverySIEOUTCOMEDateSold = parts[61];
                            if (!string.IsNullOrEmpty(parts[62]))
                            {
                                newAccfile.RecoverySIEOUTCOMESoldAmount = Convert.ToDecimal(parts[62]);
                            }
                            if (!string.IsNullOrEmpty(parts[63]))
                            {
                                newAccfile.NettCredits = Convert.ToDecimal(parts[63]);
                            }
                            if (!string.IsNullOrEmpty(parts[64]))
                            {

                                newAccfile.Arrears = Convert.ToDecimal(parts[64]);
                            }



                            if (!string.IsNullOrEmpty(parts[65]))
                            {
                                newAccfile.Balance = Convert.ToDecimal(parts[65]);
                            }
                            newAccfile.Asset = parts[66];
                            if (noFields > 67)
                            {
                                newAccfile.PrelegalInstructionReceivedDate = parts[67];
                                newAccfile.PrelegalName = parts[68];
                                newAccfile.LegalInstructionReceivedDate = parts[69];
                                newAccfile.LegalName = parts[70];
                            }
                            if (noFields > 71)
                            {


                                if (!string.IsNullOrEmpty(parts[71]))
                                {
                                    newAccfile.InstalmentAmount = Convert.ToDecimal(parts[71]);
                                }
                                newAccfile.InstalmentDueDate = parts[72];
                                newAccfile.InstalmentDate = parts[73];
                                newAccfile.DateaccountOpened = parts[74];
                                newAccfile.NID = parts[75];
                                newAccfile.NDD = parts[76];
                                newAccfile.MOB = parts[77];
                                if (!string.IsNullOrEmpty(parts[78]))
                                {
                                    newAccfile.OriginalBalance = Convert.ToDecimal(parts[78]);
                                }
                                if (!string.IsNullOrEmpty(parts[79]))
                                {

                                    newAccfile.OutstandingBalance = Convert.ToDecimal(parts[79]);
                                }
                                newAccfile.Tag6 = parts[80];
                                newAccfile.Tag6AssignmentDate = parts[81];
                                newAccfile.Tag7 = parts[82];
                                newAccfile.Tag7AssignmentDate = parts[83];
                                newAccfile.Tag8 = parts[84];
                                newAccfile.Tag8AssignmentDate = parts[85];
                                if (!string.IsNullOrEmpty(parts[86]))
                                {
                                    newAccfile.IntRate = Convert.ToDecimal(parts[86]);
                                }
                                newAccfile.LoanTerm = parts[87];
                                newAccfile.NrPTPCCC = parts[88];
                                newAccfile.NrPTPConsec = parts[89];
                                newAccfile.NrPTPLife = parts[90];
                                newAccfile.BranchName = parts[91];
                            }
                            if (noFields > 92)
                            {

                                newAccfile.RiskType = parts[92];
                                newAccfile.TRIADPool = parts[93];
                                newAccfile.CreditLifeapplicable = parts[94];
                                newAccfile.PlasticNr = parts[95];
                                newAccfile.LegalAccountNBumber = parts[96];
                                newAccfile.Unemployed = parts[97];
                                newAccfile.Unemployedsince = parts[98];
                                newAccfile.UnemploymentAffidavitDate = parts[99];
                                newAccfile.Section129Letter = parts[100];
                                newAccfile.varchargeoffDate = parts[101];
                                if (!string.IsNullOrEmpty(parts[102]))
                                {
                                    newAccfile.ArrearBalanceatvarchargeoff = Convert.ToDecimal(parts[102]);
                                }
                                if (!string.IsNullOrEmpty(parts[103]))
                                {
                                    newAccfile.OutstandingBalanceatvarchargeoff = Convert.ToDecimal(parts[103]);
                                }
                                if (!string.IsNullOrEmpty(parts[104]))
                                {
                                    newAccfile.Interestatvarchargeoff = Convert.ToDecimal(parts[104]);
                                }
                                if (!string.IsNullOrEmpty(parts[105]))
                                {

                                    newAccfile.HandoverAmount = Convert.ToDecimal(parts[105]);
                                }
                                newAccfile.AssetDescription = parts[106];
                                newAccfile.AssetRegistration = parts[107];
                                newAccfile.Enginenumber = parts[108];
                                newAccfile.Chassisnumber = parts[109];
                                newAccfile.Vinnumber = parts[110];
                                newAccfile.MMCode = parts[111];
                                newAccfile.DealerName = parts[112];
                                newAccfile.DealerNumber = parts[113];
                                if (!string.IsNullOrEmpty(parts[114]))
                                {
                                    newAccfile.TradeValue = Convert.ToDecimal(parts[114]);
                                }
                                if (!string.IsNullOrEmpty(parts[115]))
                                {

                                    newAccfile.RetailValue = Convert.ToDecimal(parts[115]);
                                }
                                if (!string.IsNullOrEmpty(parts[116]))
                                {
                                    newAccfile.SoldAmount = Convert.ToDecimal(parts[116]);
                                }
                                newAccfile.TradeCentreNumber = parts[117];
                                newAccfile.VehicleStockNumber = parts[118];
                                newAccfile.Assetinpossession = parts[119];
                                newAccfile.WarrantIssued = parts[120];
                                newAccfile.Warrant_Sheriff = parts[121].Trim();
                                newAccfile.Sheriffappointed = parts[122].Trim();
                                newAccfile.DocumentationCompleted = parts[123].Trim();
                                newAccfile.DisputeStatus = parts[124];
                                newAccfile.DisputeStatusDate = parts[125];
                                newAccfile.Complaint = parts[126];
                                newAccfile.Complaintdate = parts[127];
                                newAccfile.DebtReviewIndicator = parts[128];
                                newAccfile.InternalPlacementbefore = parts[129];
                                newAccfile.InternalPlacementdate = parts[130];
                                newAccfile.Prescriptionaccount = parts[131];
                                newAccfile.LastPTP = parts[132];
                            }
                            if (noFields > 133)
                            {
                                newAccfile.FirstName = parts[133];
                                if (!string.IsNullOrEmpty(parts[134]))
                                {
                                    newAccfile.SettlmentPercentage = Convert.ToDecimal(parts[134]);
                                }
                                if (!string.IsNullOrEmpty(parts[135]))
                                {
                                    newAccfile.PastDue = Convert.ToDecimal(parts[135]);
                                }
                                if (!string.IsNullOrEmpty(parts[136]))
                                {
                                    newAccfile.Limit = Convert.ToDecimal(parts[136]);
                                }
                                if (!string.IsNullOrEmpty(parts[137]))
                                {
                                    newAccfile.OverLimit = Convert.ToDecimal(parts[137]);
                                }
                                if (!string.IsNullOrEmpty(parts[138]))
                                {
                                    newAccfile.VAtAmt = Convert.ToDecimal(parts[138]);
                                }
                            }
                            if (noFields > 139)
                            {

                                if (!string.IsNullOrEmpty(parts[139]))
                                {
                                    newAccfile.ECBalance = Convert.ToDecimal(parts[139]);
                                }
                                if (!string.IsNullOrEmpty(parts[140]))
                                {
                                    newAccfile.CONSBalance = Convert.ToDecimal(parts[140]);
                                }
                                if (!string.IsNullOrEmpty(parts[141]))
                                {
                                    newAccfile.PrePrescribeTag = parts[141];
                                }
                            }
                            if (noFields > 142)
                            {
                                if (!string.IsNullOrEmpty(parts[142]))
                                {
                                    newAccfile.Total_OutStanding_Bal = Convert.ToDecimal(parts[142]);
                                }
                                if (!string.IsNullOrEmpty(parts[141]))
                                {
                                    newAccfile.OverDarft_Interest_Rate = Convert.ToDecimal(parts[141]);
                                }
                                if (!string.IsNullOrEmpty(parts[142]))
                                {
                                    newAccfile.Decline_Reason = parts[142];
                                }


                            }
                            if (noFields > 143)
                            {
                                if (!string.IsNullOrEmpty(parts[143]))
                                {
                                    newAccfile.Memo = parts[143];
                                }
                                if (!string.IsNullOrEmpty(parts[144]))
                                {
                                    newAccfile.XDS_ITC1= parts[144];
                                }
                                if (!string.IsNullOrEmpty(parts[145]))
                                {
                                    newAccfile.XDS_ITC2= parts[145];
                                }

                                if (!string.IsNullOrEmpty(parts[146]))
                                {
                                    newAccfile.EDC_GIVEN1 = parts[146];
                                }
                                if (!string.IsNullOrEmpty(parts[147]))
                                {
                                    newAccfile.EDC_GIVEN2 = parts[147];
                                }
                                if (!string.IsNullOrEmpty(parts[148]))
                                {
                                    newAccfile.INTERNAL_PHN1 = parts[148];
                                }
                                if (!string.IsNullOrEmpty(parts[149]))
                                {
                                    newAccfile.INTERNAL_PHN2 = parts[149];
                                }
                                if (!string.IsNullOrEmpty(parts[150]))
                                {
                                    newAccfile.INTERNAL_PHN3 = parts[150];
                                }
                                if (!string.IsNullOrEmpty(parts[151]))
                                {
                                    newAccfile.INTERNAL_PHN4 = parts[151];
                                }
                                if (!string.IsNullOrEmpty(parts[152]))
                                {
                                    newAccfile.INTERNAL_PHN5 = parts[152];
                                }
                                if (!string.IsNullOrEmpty(parts[153]))
                                {
                                    newAccfile.INTERNAL_PHN6 = parts[153];
                                }
                            }
                            if (noFields > 169)

                            {

                                if (!string.IsNullOrEmpty(parts[157]))
                                {
                                    newAccfile.Placement = parts[157];
                                }
                                if (!string.IsNullOrEmpty(parts[158]))
                                {
                                    newAccfile.Delinquency5 = parts[158];
                                }
                                if (!string.IsNullOrEmpty(parts[159]))
                                {
                                    newAccfile.Delinquency30 = parts[159];
                                }
                                if (!string.IsNullOrEmpty(parts[160]))
                                {
                                    newAccfile.Delinquency60 = parts[160];
                                }
                                if (!string.IsNullOrEmpty(parts[161]))
                                {
                                    newAccfile.Delinquency90 = parts[161];
                                }
                                if (!string.IsNullOrEmpty(parts[162]))
                                {
                                    newAccfile.Delinquency120 = parts[162];
                                }
                                if (!string.IsNullOrEmpty(parts[163]))
                                {
                                    newAccfile.Delinquency150 = parts[163];
                                }
                                if (!string.IsNullOrEmpty(parts[164]))
                                {
                                    newAccfile.Delinquency180 = parts[164];
                                }
                                if (!string.IsNullOrEmpty(parts[165]))
                                {
                                    newAccfile.Delinquency210 = parts[165];
                                }
                                if (!string.IsNullOrEmpty(parts[166]))
                                {
                                    newAccfile.Latest_Payment_Schedule_Date = parts[166];
                                }
                                if (!string.IsNullOrEmpty(parts[167]))
                                {
                                    newAccfile.OD_NONOD = parts[167];
                                }
                                if (!string.IsNullOrEmpty(parts[168]))
                                {
                                    newAccfile.OverDraftLimit = parts[168];
                                }
                                if (!string.IsNullOrEmpty(parts[169]))
                                {
                                    newAccfile.Placement_Date = parts[169];
                                }


                            }
                            recordsToBeInserted.Add(newAccfile);
                        }
                        catch (System.FormatException ex3)
                        {
                            txtAttach.Text += ex3.Message;
                            txtAttach.Text += string.Format("Error record not imported format error ") + Environment.NewLine;
                        }
                        if (noUpd >= recordsToInsert)
                        {
                            //  Insert this List<> of records using a Bulk Insert
                            DataTableHelper.BulkCopyToDatabase(recordsToBeInserted, "ABSAAccFile", dc);

                            txtAttach.Text += string.Format("Inserted  " + noUpd.ToString() + " records TO AccFiles " + allCount.ToString() ) + Environment.NewLine;
                            recordsToBeInserted.Clear();
                            noUpd = 0;

                        }

                    }
                }


                DateTime dtInsertStarted = DateTime.Now;

                //  Insert this List<> of records using a Bulk Insert
                DataTableHelper.BulkCopyToDatabase(recordsToBeInserted, "ABSAAccFile", dc);

                TimeSpan duration = DateTime.Now - dtInsertStarted;
                txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", recordsToInsert, duration.Seconds, duration.Milliseconds) + Environment.NewLine;
            }
            this.Cursor = Cursors.Arrow;
        }


        protected void ProcessPayFile(String accFileName, String NewPath)
        {
            String[] fileSplit;
            int noFields;

            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            /// string paydelStatement = "Delete from ABSAPayFile where Filename = " + accFileName.Replace(".zip", ".txt");
            /// BdSuiteContext.ExecuteQuery<ABSAAccFile>(paydelStatement);
            
             string paydelStatement = "Delete from ABSAPayFile where Filename = {0},\"" + accFileName.Replace(".zip", ".txt") + "\"";
            ///BdSuiteContext.ExecuteQuery<ABSAAccFile>(paydelStatement);
            ///BdSuiteContext.ExecuteCommand(paydelStatement);

                      
            int lineno=0;
            fileSplit = accFileName.Split('_');

            int recordsToInsert = int.Parse(tbRecordsToInsert.Text);
            this.Cursor = Cursors.WaitCursor;

            List<ABSAPayFile> recordsToBeInserted = new List<ABSAPayFile>();
            using (DataClassBDSuiteDataContext dc = new DataClassBDSuiteDataContext())
            {
                using (var textFieldParser = new TextFieldParser(NewPath + accFileName))
                {
                    textFieldParser.TextFieldType = FieldType.Delimited;
                    textFieldParser.Delimiters = new[] { "|" };
                    textFieldParser.TrimWhiteSpace = true;
                    textFieldParser.HasFieldsEnclosedInQuotes = true;
                    // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
                    while (!textFieldParser.EndOfData)
                    {
                       
                        lineno++;  
                        string[] parts = textFieldParser.ReadFields();
                        noFields = parts.Length;
                        if (noFields > 0)
                        {
                            txtProcess.Text = accFileName + " - " + noFields.ToString();

                            ABSAPayFile newPayFile = new ABSAPayFile();

                                                     String checkStr;
                            String checkData;
                            String newData;
                            checkData = parts[2];
                            checkStr = checkData;
                            CultureInfo provider = CultureInfo.InvariantCulture;

                            if (noFields > 19)
                            {

                                newPayFile.Filename = accFileName.Replace(".zip", ".txt");
                                newPayFile.AccType = fileSplit[3];
                                newPayFile.FileDate = fileSplit[4];
                                //      dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyy/MM/dd", System.Globalization.CultureInfo.InvariantCulture);
                                //      newRecallFile.DateF = Convert.ToInt32(dateNew) + 2;
                                newPayFile.Book = fileSplit[2];

                                newPayFile.Workgroup = parts[0];
                                newPayFile.AccountNumber = parts[1];
                                newPayFile.Amount = Convert.ToDecimal(parts[2]);

                                if (!string.IsNullOrEmpty(parts[3]) && parts[3].Length > 9)
                                {
                                    newData = parts[3];
                                    newPayFile.FinancialTransactionEffectiveDate = DateTime.ParseExact(parts[3], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                                    //newPayFile.FinancialTransactionEffectiveDate = Convert.ToDateTime(parts[3]);
                                    // newPayFile.FinancialTransactionEffectiveDate = DateTime.ParseExact(parts[3],"yyyy/mm/dd",provider);
                                    //DateTime dt = DateTime.ParseExact(parts[3], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                }
                                newPayFile.FileDate = fileSplit[4];
                                newPayFile.MemoCode = parts[4];
                                newPayFile.PaymentType = parts[5];
                                newPayFile.PaymentMethod = parts[6];
                               // newPayFile.Tag6 = parts[7];
                                //  Tag6AssignmentDate = Convert.ToDateTime(parts[8]),
                                newPayFile.Tag4 = parts[9];
                                if (!string.IsNullOrEmpty(parts[10]))
                                {
                                    // newPayFile.Tag4AssignmentDate = Convert.ToDateTime(parts[10]);
                                    newPayFile.Tag4AssignmentDate = DateTime.ParseExact(parts[10], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                                }
                                newPayFile.Tag5 = parts[11];
                                if (!string.IsNullOrEmpty(parts[12]))
                                {
                                    //newPayFile.Tag5AssignmentDate = Convert.ToDateTime(parts[12]);
                                    newPayFile.Tag5AssignmentDate = DateTime.ParseExact(parts[12], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                                }
                                newPayFile.NAEDO = parts[13];
                                if (!string.IsNullOrEmpty(parts[14]))
                                {
                                    //newPayFile.Dateaccountopened = Convert.ToDateTime(parts[14]);
                                    newPayFile.Dateaccountopened = DateTime.ParseExact(parts[14], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                                }
                                newPayFile.SiteName = parts[15];
                                newPayFile.IDNumber = parts[16];
                                newPayFile.Surname = parts[17];
                                newPayFile.SBUCode = parts[18];
                                newPayFile.Site = parts[19];
                            }
                            if (noFields > 20)
                            {
                                newPayFile.AttorneyCode = parts[20];
                                newPayFile.Commissionable = parts[21];
                            }

                            if (noFields > 28)
                            {
                                newPayFile.Placement_Tag1 = parts[22];
                                newPayFile.Placement_Tag2 = parts[23];
                                newPayFile.NID_NDD = parts[24];
                                newPayFile.Business_Area= parts[25];
                                newPayFile.OD_NONOD = parts[26];
                                newPayFile.Overdraft_limit = parts[27];
                                newPayFile.Product = parts[28];

                            }


                            recordsToBeInserted.Add(newPayFile);
                        }
                    }
                }


                DateTime dtInsertStarted = DateTime.Now;

                //  Insert this List<> of records using a Bulk Insert
                DataTableHelper.BulkCopyToDatabase(recordsToBeInserted, "ABSAPayFile", dc);

                TimeSpan duration = DateTime.Now - dtInsertStarted;
                txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", recordsToInsert, duration.Seconds, duration.Milliseconds) + Environment.NewLine;
            }
            this.Cursor = Cursors.Arrow;
        }

        protected void ProcessRecallFile(String accFileName, String NewPath)
        {
            String[] fileSplit;
            int noFields;
           
            fileSplit = accFileName.Split('_');

            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            ///   string recalldelStatement = "Delete from ABSARecallFile where Filename = " + accFileName.Replace(".zip", ".txt");
            ///   BdSuiteContext.ExecuteQuery<ABSAAccFile>(recalldelStatement);
            ///   
               string recalldelStatement = "Delete from ABSARecallFile where Filename = {0},\"" + accFileName.Replace(".zip", ".txt") + "\"";
             ///  BdSuiteContext.ExecuteQuery<ABSAAccFile>(recalldelStatement);
            ///BdSuiteContext.ExecuteCommand(recalldelStatement);

            int recordsToInsert = int.Parse(tbRecordsToInsert.Text);
            this.Cursor = Cursors.WaitCursor;

            List<ABSARecallFile> recordsToBeInserted = new List<ABSARecallFile>();
            using (DataClassBDSuiteDataContext dc = new DataClassBDSuiteDataContext())
            {
                using (var textFieldParser = new TextFieldParser(NewPath + accFileName))
                {
                    textFieldParser.TextFieldType = FieldType.Delimited;
                    textFieldParser.Delimiters = new[] { "|" };
                    textFieldParser.TrimWhiteSpace = true;
                    textFieldParser.HasFieldsEnclosedInQuotes = true;
                    // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
                    while (!textFieldParser.EndOfData)
                    {
                        string[] parts = textFieldParser.ReadFields();
                        noFields = parts.Length;

                        if (noFields > 0)
                        {
                            ABSARecallFile newRecallFile = new ABSARecallFile();
                            newRecallFile.Filename = accFileName.Replace(".zip", ".txt");
                            newRecallFile.AccType = fileSplit[3];
                            newRecallFile.FileDate = fileSplit[4];
                            //      dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyy/MM/dd", System.Globalization.CultureInfo.InvariantCulture);
                            //      newRecallFile.DateF = Convert.ToInt32(dateNew) + 2;
                            newRecallFile.Book = fileSplit[2];
                            newRecallFile.EDCWorkGroup = parts[0];
                            newRecallFile.Accountnumber = parts[1];
                            newRecallFile.LegalAccountNumber = parts[2];
                            newRecallFile.PrimaryPlasticNumber = parts[3];
                            DateTime dt = DateTime.ParseExact(parts[4], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                            newRecallFile.RecallDate = dt.Date;
                            //newRecallFile.RecallDate = Convert.ToDateTime(parts[4]);
                            if (noFields > 5)
                            {
                                newRecallFile.Recall_Area = parts[5];
                                newRecallFile.Prev_Risk = parts[6];
                                newRecallFile.Current_Risk = parts[7];
                            }

                            recordsToBeInserted.Add(newRecallFile);
                        }
                    }
                }


                DateTime dtInsertStarted = DateTime.Now;

                //  Insert this List<> of records using a Bulk Insert
                DataTableHelper.BulkCopyToDatabase(recordsToBeInserted, "ABSARecallFile", dc);

                TimeSpan duration = DateTime.Now - dtInsertStarted;
                txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", recordsToInsert, duration.Seconds, duration.Milliseconds) + Environment.NewLine;
            }
            this.Cursor = Cursors.Arrow;
        }

        public void ConvertPipedCSVToDataTable()
        {
            try
            {

                //Declare Variable and set values
                //provide input folder
                string SourceFolder = @"C:\Source\";
                //provide file name
                string SourceFileName = "TechBrothersIT.txt";
                //provide the table name in which you would like to load data
                string TableName = "dbo.TechBrothersITCourse";
                //provide the file delimiter such as comma,pipe
                string filedelimiter = ",";
                //provide the Archive folder where you would like to move file
                string ArchiveFodler = @"C:\Archive\";

                //Create Connection to SQL Server
                SqlConnection SQLConnection = new SqlConnection();
                SQLConnection.ConnectionString = "Data Source = (local); Initial Catalog =TechBrothersIT; "
                                 + "Integrated Security=true;";

                System.IO.StreamReader SourceFile =
                new System.IO.StreamReader(SourceFolder + SourceFileName);

                string line = "";
                Int32 counter = 0;

                SQLConnection.Open();
                while ((line = SourceFile.ReadLine()) != null)
                {
                    //skip the header row
                    if (counter > 0)
                    {
                        //prepare insert query
                        string query = "Insert into " + TableName +
                               " Values ('" + line.Replace(filedelimiter, "'|'") + "')";

                        //execute sqlcommand to insert record
                        SqlCommand myCommand = new SqlCommand(query, SQLConnection);
                        myCommand.ExecuteNonQuery();
                    }
                    counter++;
                }

                SourceFile.Close();
                SQLConnection.Close();

                //Move the file to Archive folder
                File.Move(SourceFolder + SourceFileName, ArchiveFodler + SourceFileName);

            }
            catch (IOException Exception)
            {
                Console.Write(Exception);
            }
        }

                protected const string _truncateLiveTableCommandText = @"TRUNCATE TABLE YourTableName";
        protected const int _batchSize = 100000;

        public void LoadCsvDataIntoSqlServer()
        {
            // This should be the full path
            var fileName = @"C:\Path\To\File.csv";

            var createdCount = 0;

            using (var textFieldParser = new TextFieldParser(fileName))
            {
                textFieldParser.TextFieldType = FieldType.Delimited;
                textFieldParser.Delimiters = new[] { "|" };
                textFieldParser.HasFieldsEnclosedInQuotes = true;

                var connectionString = ConfigurationManager.ConnectionStrings["BDSuiteConnectionString"].ConnectionString;

                var dataTable = new DataTable("YourTableName");

                // Add the columns in the temp table
                dataTable.Columns.Add("FirstName");
                dataTable.Columns.Add("LastName");

                using (var sqlConnection = new SqlConnection(connectionString))
                {
                    sqlConnection.Open();

                    // Truncate the live table
                    using (var sqlCommand = new SqlCommand(_truncateLiveTableCommandText, sqlConnection))
                    {
                        sqlCommand.ExecuteNonQuery();
                    }

                    // Create the bulk copy object
                    var sqlBulkCopy = new SqlBulkCopy(sqlConnection)
                    {
                        DestinationTableName = "YourTableName"
                    };

                    // Setup the column mappings, anything ommitted is skipped
                    sqlBulkCopy.ColumnMappings.Add("FirstName", "FirstName");
                    sqlBulkCopy.ColumnMappings.Add("LastName", "LastName");

                    // Loop through the CSV and load each set of 100,000 records into a DataTable
                    // Then send it to the LiveTable
                    while (!textFieldParser.EndOfData)
                    {
                        dataTable.Rows.Add(textFieldParser.ReadFields());

                        createdCount++;

                        if (createdCount % _batchSize == 0)
                        {
                            InsertDataTable(sqlBulkCopy, sqlConnection, dataTable);

                        }
                    }

                    // Don't forget to send the last batch under 100,000
                    InsertDataTable(sqlBulkCopy, sqlConnection, dataTable);

                    sqlConnection.Close();
                }
            }
        }

        protected void InsertDataTable(SqlBulkCopy sqlBulkCopy, SqlConnection sqlConnection, DataTable dataTable)
        {
            sqlBulkCopy.WriteToServer(dataTable);

            dataTable.Rows.Clear();
        }

        private void lstBxAccount_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtlstBxAccount.Text= lstBxAccount.SelectedItem.ToString();
        }

        private void btnGetProcess_Click(object sender, EventArgs e)
        {
            lstBxAccount.Items.Clear();
            nList.Clear();
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            var q = from p in BdSuiteContext.AbsaFileLogs
                    where p.Processed == 0
                    orderby p.FileDate ascending, p.Book ascending
                    select p;
            foreach(var log in q )
            {

                lstBxAccount.Items.Add(log.Filename);
                nList.Add(log.Filename);

            }
            MessageBox.Show("Completed Getting Unprocessed File");
        }

        private void btnProcessAllFiles_Click(object sender, EventArgs e)
        {
           
            int i=1;
            string[] fileSplit;
            string[] extensionsArray = { "_PAYMENT", "_RECALL", "_ACCOUNT","_" };
            foreach  (var itemAcc in nList)
                   {
                var fn = itemAcc.ToString();
                if (extensionsArray.Any(fn.Contains))
                {
                    DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
                    txtAttach.Text += itemAcc.ToString() + Environment.NewLine;
                    if (File.Exists(servPath + "ABSATextFile\\" + itemAcc.ToString()))
                    {
                        ProcessFiles(itemAcc.ToString(), servPath + "ABSATextFile\\");

                        fileSplit = itemAcc.Split('_');
                        var q = from p in BdSuiteContext.AbsaFileLogs
                                where p.Filename == itemAcc.Replace(".zip", ".txt")
                                select p;
                        if (q.Count() == 0) 
                        

                        {
                            AbsaFileLog objFilelog = new AbsaFileLog();
                            objFilelog.Filename = itemAcc;
                            objFilelog.AccType = "ACCOUNT";
                            objFilelog.FileDate = fileSplit[3];
                            objFilelog.Book = fileSplit[2];
                            objFilelog.CompPath = basePath;
                            objFilelog.ServPath = servPath;
                            objFilelog.Processed = 1;
                            BdSuiteContext.AbsaFileLogs.InsertOnSubmit(objFilelog);
                            BdSuiteContext.SubmitChanges();
                        }
                        else          
                        {

                            var r = (from p in BdSuiteContext.AbsaFileLogs
                                     where p.Filename == itemAcc
                                     select p).SingleOrDefault();
                           
                                r.Processed = 1;
                            BdSuiteContext.SubmitChanges();
                        }

                        if (i <= lstBxAccount.Items.Count - 1)
                        {
                            lstBxAccount.SelectedIndex = i;
                            i++;
                        }
                    }
                    else
                    {
                        var q = (from p in BdSuiteContext.AbsaFileLogs
                                 where p.Filename == itemAcc
                                 select p).SingleOrDefault();
                        q.Processed = 1;
                        BdSuiteContext.SubmitChanges();
                    }
                }

            }
            MessageBox.Show("Completed Processing File");

        }


        private void btnProcessOneFiles_Click(object sender, EventArgs e)
        {

        }


        protected void ProcessFilesLen(String processFileName, String NewPath)

        {
            {
                ProcessAllFileLen(processFileName, NewPath);
            }

        }


        protected void ProcessAllFileLen(String accFileName, String NewPath)
        {
            int noFields;
            string allData;
            String[] fileSplit;
          

            fileSplit = accFileName.Split('_');
            using (var textFieldParser = new TextFieldParser(NewPath + accFileName))
            {
                textFieldParser.TextFieldType = FieldType.Delimited;
                textFieldParser.Delimiters = new[] { "|" };
                textFieldParser.TrimWhiteSpace = true;
                textFieldParser.HasFieldsEnclosedInQuotes = true;
                // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
                if (!textFieldParser.EndOfData)
                {
                    allData = accFileName + "," + fileSplit[3];
                    string[] parts = textFieldParser.ReadFields();
                    noFields = parts.Length;
                    allData = accFileName + "," + fileSplit[3] + "," + noFields.ToString() + ",";
                    txtAttach.Text += "No fields " + noFields.ToString() + Environment.NewLine;
                    foreach (var dataLine in parts)
                    {
                        allData = allData + "," + dataLine.Length.ToString();

                    }
                }
                else
                {
                    allData = accFileName + ",0 - No Record";
                }

                txtAttach.Text += allData + Environment.NewLine;
            }
        }

        private void btnFileReProcess_Click(object sender, EventArgs e)
        {
            lstBxAccount.Items.Clear();
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            var q = from p in BdSuiteContext.AbsaFileLogs
                    where p.AccType == "ACCOUNT"  && p.FileDate.StartsWith("2018")
                    orderby p.DateF descending, p.Book descending
                    select p;
            foreach (var log in q)
            {

                lstBxAccount.Items.Add(log.Filename);
                nList.Add(log.Filename);

            }
        }

        private void btnCheckFiles_Click(object sender, EventArgs e)
        {
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            int i = 1;
            string[] extensionsArray = { "_PAYMENT", "_RECALL", "_ACCOUNT" };
            foreach (var itemAcc in nList)
            {
                var fn = itemAcc.ToString();
                if (extensionsArray.Any(fn.Contains))
                {
                    txtAttach.Text += itemAcc.ToString() + Environment.NewLine;
                    if (File.Exists(basePath + itemAcc.ToString()))
                    {
                        CheckProcessFiles(itemAcc.ToString(), basePath);

                    if (i <= lstBxAccount.Items.Count - 1)
                        {
                            lstBxAccount.SelectedIndex = i;
                            i++;
                        }
                    }
                    else
                    {
                       
                    }
                }

            }
            MessageBox.Show("Completed Processing File");

        }

        protected void CheckProcessFiles(String processFileName, String NewPath)

        {
            String[] fileSplit;

            fileSplit = processFileName.Split('_');
            if (fileSplit.Length > 2)
            {
                if (fileSplit[3] == "ACCOUNT")
                {
                    CheckProcessAccFile(processFileName, NewPath);
                }
                if (fileSplit[3] == "RECALL")
                {
                    CheckProcessRecallFile(processFileName, NewPath);
                }
                if (fileSplit[3] == "PAYMENT")
                {
                    CheckProcessPayFile(processFileName, NewPath);
                }
            }
        }
        protected void CheckProcessAccFile(String accFileName, String NewPath)
        {
            String[] fileSplit;
            int noFields;
            int noUpd = 0;
            int allCount = 0;
            String newdata;

            DateTime dateNew;

            fileSplit = accFileName.Split('_');

                    var textFieldParser = new TextFieldParser(NewPath + accFileName);
                
                    textFieldParser.TextFieldType = FieldType.Delimited;
                    textFieldParser.Delimiters = new[] { "|" };
                    textFieldParser.TrimWhiteSpace = true;
                    textFieldParser.HasFieldsEnclosedInQuotes = true;
                    // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
                    while (!textFieldParser.EndOfData)
                    {
                        try
                        {
                            noUpd++;
                            allCount++;
                            string[] parts = textFieldParser.ReadFields();
                            noFields = parts.Length;
                            txtProcess.Text = accFileName + " - " + noFields.ToString();
                            ABSAAccFile newAccfile = new ABSAAccFile();

                            newAccfile.Filename = accFileName.Replace(".zip", ".txt");
                            newAccfile.AccType = fileSplit[3];
                            newAccfile.FileDate = fileSplit[4];
                            dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                            //    newAccfile.DateF = Convert.tol(dateNew) + 2;
                            newAccfile.Book = fileSplit[2];
                            newAccfile.Ind = parts[0];
                            newAccfile.Accountnumber = parts[1];
                            newAccfile.Title = parts[2];
                            newAccfile.Initials = parts[3];
                            newAccfile.Surname = parts[4];
                            newAccfile.IDNumber = parts[5];
                            newAccfile.EmailAddress = parts[6];
                            newAccfile.CellNo = parts[7];
                            newAccfile.HomeTel = parts[8];
                            newAccfile.WorkTel = parts[9];
                            newAccfile.FaxNo = parts[10];
                            newAccfile.HomeLine1 = parts[11];
                            newAccfile.HomeLine2 = parts[12];
                            newAccfile.HomeLine3 = parts[13];
                            newAccfile.HomeLine4 = parts[14];
                            newAccfile.HomePostal = parts[15];
                            newAccfile.PostalLine1 = parts[16];
                            newAccfile.PostalLine2 = parts[17];
                            newAccfile.PostalLine3 = parts[18];
                            newAccfile.PostalLine4 = parts[19];
                            newAccfile.PostalCode = parts[20];
                            newAccfile.EmployerName = parts[21];
                            newAccfile.EmployerAddLine1 = parts[22];
                            newAccfile.EmployerAddLine2 = parts[23];
                            newAccfile.EmployerAddLine3 = parts[24];
                            newAccfile.EmployerPostalCode = parts[25];
                            newAccfile.Occupation = parts[26];
                            newAccfile.Product = parts[27];
                            newAccfile.InstructionReceivedDate = parts[28];
                            newAccfile.InterestFromDate = parts[29];
                            newdata = parts[30];
                            if (!string.IsNullOrEmpty(parts[30]) || parts[30] == "0.00")
                            {
                                newAccfile.InterestRate = Convert.ToDecimal(parts[30]);
                            }
                            newAccfile.SBUCode = parts[31];
                            newAccfile.BranchNr = parts[32];
                            newAccfile.LastPaymentDate = parts[33];
                            if (!string.IsNullOrEmpty(parts[34]) || parts[34] == "0.00")
                            {

                                newAccfile.LastPaymentAmount = Convert.ToDecimal(parts[34]);
                            }
                            newAccfile.Tag1 = parts[35];
                            newAccfile.Tagassignmentdate1 = parts[36];
                            newAccfile.Tag2 = parts[37];
                            newAccfile.Tagassignmentdate2 = parts[38];
                            newAccfile.Tag3 = parts[39];
                            newAccfile.Tagassignmentdate3 = parts[40];
                            newAccfile.Tag4 = parts[41];
                            newAccfile.Tagassignmentdate4 = parts[42];
                            newAccfile.Tag5 = parts[43];
                            newAccfile.Tagassignmentdate5 = parts[44];
                            newAccfile.SummonsCourt = parts[45];
                            newAccfile.SummonsCaseNumber = parts[46];
                            newAccfile.SummonsServedDate = parts[47];
                            newAccfile.JudgementCaseNumber = parts[48];
                            newAccfile.JudgementCourtDate = parts[49];
                            newAccfile.LegalJudgementDAte = parts[50];
                            if (!string.IsNullOrEmpty(parts[51]))
                            {
                                newAccfile.DefaultJudgementAmount = Convert.ToDecimal(parts[51]);
                            }
                            newAccfile.SIEOUTCOMEDateSold = parts[52];
                            if (!string.IsNullOrEmpty(parts[53]))
                            {

                                newAccfile.SIEOUTCOMESoldAmount = Convert.ToDecimal(parts[53]);
                            }
                            newAccfile.DateWrittenOff = parts[54];
                            if (!string.IsNullOrEmpty(parts[55]))
                            {
                                newAccfile.CapitalWrittenOff = Convert.ToDecimal(parts[55]);
                            }
                            if (!string.IsNullOrEmpty(parts[56]))
                            {
                                newAccfile.InterestWrittenOff = Convert.ToDecimal(parts[56]);
                            }
                            newAccfile.RecoverySummonsCourt = parts[57];
                            newAccfile.RecoverySummonsCase = parts[58];
                            newAccfile.RecoverySummonsServedDate = parts[59];
                            newAccfile.RecoveryJudgementDate = parts[60];
                            newAccfile.RecoverySIEOUTCOMEDateSold = parts[61];
                            if (!string.IsNullOrEmpty(parts[62]))
                            {
                                newAccfile.RecoverySIEOUTCOMESoldAmount = Convert.ToDecimal(parts[62]);
                            }
                            if (!string.IsNullOrEmpty(parts[63]))
                            {
                                newAccfile.NettCredits = Convert.ToDecimal(parts[63]);
                            }
                            if (!string.IsNullOrEmpty(parts[64]))
                            {

                                newAccfile.Arrears = Convert.ToDecimal(parts[64]);
                            }



                            if (!string.IsNullOrEmpty(parts[65]))
                            {
                                newAccfile.Balance = Convert.ToDecimal(parts[65]);
                            }
                            newAccfile.Asset = parts[66];
                            if (noFields > 67)
                            {
                                newAccfile.PrelegalInstructionReceivedDate = parts[67];
                                newAccfile.PrelegalName = parts[68];
                                newAccfile.LegalInstructionReceivedDate = parts[69];
                                newAccfile.LegalName = parts[70];
                            }
                            if (noFields > 71)
                            {


                                if (!string.IsNullOrEmpty(parts[71]))
                                {
                                    newAccfile.InstalmentAmount = Convert.ToDecimal(parts[71]);
                                }
                                newAccfile.InstalmentDueDate = parts[72];
                                newAccfile.InstalmentDate = parts[73];
                                newAccfile.DateaccountOpened = parts[74];
                                newAccfile.NID = parts[75];
                                newAccfile.NDD = parts[76];
                                newAccfile.MOB = parts[77];
                                if (!string.IsNullOrEmpty(parts[78]))
                                {
                                    newAccfile.OriginalBalance = Convert.ToDecimal(parts[78]);
                                }
                                if (!string.IsNullOrEmpty(parts[79]))
                                {

                                    newAccfile.OutstandingBalance = Convert.ToDecimal(parts[79]);
                                }
                                newAccfile.Tag6 = parts[80];
                                newAccfile.Tag6AssignmentDate = parts[81];
                                newAccfile.Tag7 = parts[82];
                                newAccfile.Tag7AssignmentDate = parts[83];
                                newAccfile.Tag8 = parts[84];
                                newAccfile.Tag8AssignmentDate = parts[85];
                                if (!string.IsNullOrEmpty(parts[86]))
                                {
                                    newAccfile.IntRate = Convert.ToDecimal(parts[86]);
                                }
                                newAccfile.LoanTerm = parts[87];
                                newAccfile.NrPTPCCC = parts[88];
                                newAccfile.NrPTPConsec = parts[89];
                                newAccfile.NrPTPLife = parts[90];
                                newAccfile.BranchName = parts[91];
                            }
                            if (noFields > 92)
                            {

                                newAccfile.RiskType = parts[92];
                                newAccfile.TRIADPool = parts[93];
                                newAccfile.CreditLifeapplicable = parts[94];
                                newAccfile.PlasticNr = parts[95];
                                newAccfile.LegalAccountNBumber = parts[96];
                                newAccfile.Unemployed = parts[97];
                                newAccfile.Unemployedsince = parts[98];
                                newAccfile.UnemploymentAffidavitDate = parts[99];
                                newAccfile.Section129Letter = parts[100];
                                newAccfile.varchargeoffDate = parts[101];
                                if (!string.IsNullOrEmpty(parts[102]))
                                {
                                    newAccfile.ArrearBalanceatvarchargeoff = Convert.ToDecimal(parts[102]);
                                }
                                if (!string.IsNullOrEmpty(parts[103]))
                                {
                                    newAccfile.OutstandingBalanceatvarchargeoff = Convert.ToDecimal(parts[103]);
                                }
                                if (!string.IsNullOrEmpty(parts[104]))
                                {
                                    newAccfile.Interestatvarchargeoff = Convert.ToDecimal(parts[104]);
                                }
                                if (!string.IsNullOrEmpty(parts[105]))
                                {

                                    newAccfile.HandoverAmount = Convert.ToDecimal(parts[105]);
                                }
                                newAccfile.AssetDescription = parts[106];
                                newAccfile.AssetRegistration = parts[107];
                                newAccfile.Enginenumber = parts[108];
                                newAccfile.Chassisnumber = parts[109];
                                newAccfile.Vinnumber = parts[110];
                                newAccfile.MMCode = parts[111];
                                newAccfile.DealerName = parts[112];
                                newAccfile.DealerNumber = parts[113];
                                if (!string.IsNullOrEmpty(parts[114]))
                                {
                                    newAccfile.TradeValue = Convert.ToDecimal(parts[114]);
                                }
                                if (!string.IsNullOrEmpty(parts[115]))
                                {

                                    newAccfile.RetailValue = Convert.ToDecimal(parts[115]);
                                }
                                if (!string.IsNullOrEmpty(parts[116]))
                                {
                                    newAccfile.SoldAmount = Convert.ToDecimal(parts[116]);
                                }
                                newAccfile.TradeCentreNumber = parts[117];
                                newAccfile.VehicleStockNumber = parts[118];
                                newAccfile.Assetinpossession = parts[119];
                                newAccfile.WarrantIssued = parts[120];
                                newAccfile.Warrant_Sheriff = parts[121].Trim();
                                newAccfile.Sheriffappointed = parts[122].Trim();
                                newAccfile.DocumentationCompleted = parts[123].Trim();
                                newAccfile.DisputeStatus = parts[124];
                                newAccfile.DisputeStatusDate = parts[125];
                                newAccfile.Complaint = parts[126];
                                newAccfile.Complaintdate = parts[127];
                                newAccfile.DebtReviewIndicator = parts[128];
                                newAccfile.InternalPlacementbefore = parts[129];
                                newAccfile.InternalPlacementdate = parts[130];
                                newAccfile.Prescriptionaccount = parts[131];
                                newAccfile.LastPTP = parts[132];
                            }
                            if (noFields > 133)
                            {
                                newAccfile.FirstName = parts[133];
                                if (!string.IsNullOrEmpty(parts[134]))
                                {
                                    newAccfile.SettlmentPercentage = Convert.ToDecimal(parts[134]);
                                }
                                if (!string.IsNullOrEmpty(parts[135]))
                                {
                                    newAccfile.PastDue = Convert.ToDecimal(parts[135]);
                                }
                                if (!string.IsNullOrEmpty(parts[136]))
                                {
                                    newAccfile.Limit = Convert.ToDecimal(parts[136]);
                                }
                                if (!string.IsNullOrEmpty(parts[137]))
                                {
                                    newAccfile.OverLimit = Convert.ToDecimal(parts[137]);
                                }
                                if (!string.IsNullOrEmpty(parts[138]))
                                {
                                    newAccfile.VAtAmt = Convert.ToDecimal(parts[138]);
                                }
                            }
                            if (noFields > 139)
                            {

                                if (!string.IsNullOrEmpty(parts[139]))
                                {
                                    newAccfile.ECBalance = Convert.ToDecimal(parts[139]);
                                }
                                if (!string.IsNullOrEmpty(parts[140]))
                                {
                                    newAccfile.CONSBalance = Convert.ToDecimal(parts[140]);
                                }
                                if (!string.IsNullOrEmpty(parts[141]))
                                {
                                    newAccfile.PrePrescribeTag = parts[141];
                                }
                            }
                            if (noFields > 142)
                            {
                                if (!string.IsNullOrEmpty(parts[142]))
                                {
                                    newAccfile.Total_OutStanding_Bal = Convert.ToDecimal(parts[142]);
                                }
                                if (!string.IsNullOrEmpty(parts[143]))
                                {
                                    newAccfile.OverDarft_Interest_Rate = Convert.ToDecimal(parts[143]);
                                }
                                if (!string.IsNullOrEmpty(parts[144]))
                                {
                                    newAccfile.Prev_Install_Amount = Convert.ToDecimal(parts[143]);
                                }
                                if (!string.IsNullOrEmpty(parts[145]))
                                {
                                    newAccfile.Decline_Reason = (parts[145]);
                                }
                                if (!string.IsNullOrEmpty(parts[146]))
                                {
                                    newAccfile.Memo =(parts[146]);
                                }
                                if (!string.IsNullOrEmpty(parts[147]))
                                {
                                    newAccfile.XDS_ITC1 = parts[147];
                                }

                                if (!string.IsNullOrEmpty(parts[148]))
                                {
                                    newAccfile.XDS_ITC2 = (parts[148]);
                                }
                                if (!string.IsNullOrEmpty(parts[149]))
                                {
                                    newAccfile.EDC_GIVEN1 = (parts[149]);
                                }
                                if (!string.IsNullOrEmpty(parts[150]))
                                {
                                    newAccfile.EDC_GIVEN2 = parts[150];
                                }
                                if (!string.IsNullOrEmpty(parts[151]))
                                {
                                    newAccfile.INTERNAL_PHN1 =(parts[151]);
                                }
                                if (!string.IsNullOrEmpty(parts[152]))
                                {
                                    newAccfile.INTERNAL_PHN2 = (parts[152]);
                                }
                                if (!string.IsNullOrEmpty(parts[153]))
                                {
                                    newAccfile.INTERNAL_PHN3 = parts[153];
                                }
                                if (!string.IsNullOrEmpty(parts[154]))
                                {
                                    newAccfile.INTERNAL_PHN4 = (parts[154]);
                                }
                                if (!string.IsNullOrEmpty(parts[155]))
                                {
                                    newAccfile.INTERNAL_PHN5 = (parts[155]);
                                }
                                if (!string.IsNullOrEmpty(parts[156]))
                                {
                                    newAccfile.INTERNAL_PHN6 = parts[156];
                                }
                            }
                            if (noFields > 157)
                            {
                                if (!string.IsNullOrEmpty(parts[157]))
                                {
                                    newAccfile.Placement = (parts[157]);
                                }

                                if (!string.IsNullOrEmpty(parts[158]))
                                {
                                    newAccfile.Delinquency5 = (parts[158]);
                                }
                                if (!string.IsNullOrEmpty(parts[159]))
                                {
                                    newAccfile.Delinquency30 = (parts[159]);
                                }
                                if (!string.IsNullOrEmpty(parts[160]))
                                {
                                    newAccfile.Delinquency60 = parts[160];
                                }
                                if (!string.IsNullOrEmpty(parts[161]))
                                {
                                    newAccfile.Delinquency90 = (parts[161]);
                                }
                                if (!string.IsNullOrEmpty(parts[162]))
                                {
                                    newAccfile.Delinquency120 = (parts[162]);
                                }
                                if (!string.IsNullOrEmpty(parts[163]))
                                {
                                    newAccfile.Delinquency150 = parts[163];
                                }
                                if (!string.IsNullOrEmpty(parts[164]))
                                {
                                    newAccfile.Delinquency180 = parts[164];
                                }
                                if (!string.IsNullOrEmpty(parts[165]))
                                {
                                    newAccfile.Delinquency210 = (parts[165]);
                                }
                                if (!string.IsNullOrEmpty(parts[166]))
                                {
                                    newAccfile.Latest_Payment_Schedule_Date = (parts[166]);
                                }
                                if (!string.IsNullOrEmpty(parts[167]))
                                {
                                    newAccfile.OD_NONOD = parts[167];
                                }
                                if (!string.IsNullOrEmpty(parts[168]))
                                {
                                    newAccfile.OverDraftLimit = parts[168];
                                }
                                if (!string.IsNullOrEmpty(parts[169]))
                                {
                                    newAccfile.Placement_Date = (parts[169]);
                                }
                                

                    }




                }
                        catch (System.FormatException ex3)
                        {
                            txtAttach.Text += ex3.Message;
                            txtAttach.Text += string.Format("Error record not imported format error ") + Environment.NewLine;
                        }
                    }
                


                DateTime dtInsertStarted = DateTime.Now;
                TimeSpan duration = DateTime.Now - dtInsertStarted;
                txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds","", duration.Seconds, duration.Milliseconds) + Environment.NewLine;
            
            this.Cursor = Cursors.Arrow;
        }

        protected void CheckProcessPayFile(String accFileName, String NewPath)
        {
            String[] fileSplit;
            int noFields;
            CultureInfo provider = CultureInfo.InvariantCulture;
            DateTime dateNew;

            int lineno = 0;

            fileSplit = accFileName.Split('_');
            String checkStr;
            String checkData;

            var textFieldParser = new TextFieldParser(NewPath + accFileName);
                
                    textFieldParser.TextFieldType = FieldType.Delimited;
                    textFieldParser.Delimiters = new[] { "|" };
                    textFieldParser.TrimWhiteSpace = true;
                    textFieldParser.HasFieldsEnclosedInQuotes = true;
                    // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
                    while (!textFieldParser.EndOfData)
                    {

                        lineno++;
                        string[] parts = textFieldParser.ReadFields();
                        noFields = parts.Length;

                        txtProcess.Text = accFileName + " - " + noFields.ToString();

                        ABSAPayFile newPayFile = new ABSAPayFile();

                        newPayFile.Filename = accFileName.Replace(".zip", ".txt");
                        newPayFile.AccType = fileSplit[3];
                        dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        //   newPayFile.DateF = Convert.ToInt32(dateNew) + 2;
                        newPayFile.FileDate = fileSplit[4];
                        newPayFile.Book = fileSplit[2];

                        checkData = parts[2];
                        checkStr = checkData;

                if (noFields < 19)
                {

                    newPayFile.AccountNumber = parts[0];
                    newPayFile.Amount = Convert.ToDecimal(parts[1]);
                    if (!string.IsNullOrEmpty(parts[2]) && parts[2].Length > 9)
                    {
                        //  newPayFile.FinancialTransactionEffectiveDate = DateTime.ParseExact(parts[2], "yyyy/MM/dd", provider);
                    }
                    newPayFile.MemoCode = parts[3];
                    newPayFile.PaymentType = parts[4];
                    newPayFile.PaymentMethod = parts[5];
                }
                if (noFields >19)
                { 


                            newPayFile.Workgroup = parts[0];
                            newPayFile.AccountNumber = parts[1];
                            newPayFile.Amount = Convert.ToDecimal(parts[2]);
                            if (!string.IsNullOrEmpty(parts[3]))
                            {
                                newPayFile.FinancialTransactionEffectiveDate = Convert.ToDateTime(parts[3]);
                            }
                            newPayFile.MemoCode = parts[4];
                            newPayFile.PaymentType = parts[5];
                            newPayFile.PaymentMethod = parts[6];
                            // Tag6 = parts[7],
                            //  Tag6AssignmentDate = Convert.ToDateTime(parts[8]),
                            newPayFile.Tag4 = parts[9];
                            if (!string.IsNullOrEmpty(parts[10]))
                            {
                                newPayFile.Tag4AssignmentDate = Convert.ToDateTime(parts[10]);
                            }
                            newPayFile.Tag5 = parts[11];
                            if (!string.IsNullOrEmpty(parts[12]))
                            {
                                newPayFile.Tag5AssignmentDate = Convert.ToDateTime(parts[12]);
                            }
                            newPayFile.NAEDO = parts[13];
                            if (!string.IsNullOrEmpty(parts[14]))
                            {
                                newPayFile.Dateaccountopened = Convert.ToDateTime(parts[14]);
                            }
                            newPayFile.SiteName = parts[15];
                            newPayFile.IDNumber = parts[16];
                            newPayFile.Surname = parts[17];
                            newPayFile.SBUCode = parts[18];
                            newPayFile.Site = parts[19];
                        }
                        if (noFields >= 21)
                        {
                            newPayFile.AttorneyCode = parts[20];
                            newPayFile.Commissionable = parts[21];
                        }
                        if (noFields >= 22)
                        {
                            newPayFile.Placement_Tag1 = parts[22];
                            newPayFile.Placement_Tag2 = parts[23];
                            newPayFile.NID_NDD = parts[24];
                            newPayFile.Business_Area = parts[25];
                            newPayFile.OD_NONOD = parts[26];
                            newPayFile.Overdraft_limit = parts[27];
                            newPayFile.Product = parts[28];
                            
                         }

            }
                


                DateTime dtInsertStarted = DateTime.Now;

                  TimeSpan duration = DateTime.Now - dtInsertStarted;
                txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", "", duration.Seconds, duration.Milliseconds) + Environment.NewLine;
            
            this.Cursor = Cursors.Arrow;
        }

        protected void CheckProcessRecallFile(String accFileName, String NewPath)
        {
            String[] fileSplit;
            CultureInfo provider = CultureInfo.InvariantCulture;
            DateTime dateNew;
            fileSplit = accFileName.Split('_');
            int noFields;
            var textFieldParser = new TextFieldParser(NewPath + accFileName);
                
                    textFieldParser.TextFieldType = FieldType.Delimited;
                    textFieldParser.Delimiters = new[] { "|" };
                    textFieldParser.TrimWhiteSpace = true;
                    textFieldParser.HasFieldsEnclosedInQuotes = true;
                    // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
                    while (!textFieldParser.EndOfData)
                    {
                        string[] parts = textFieldParser.ReadFields();
                noFields = parts.Length;
                ABSARecallFile newRecallFile = new ABSARecallFile();

                        newRecallFile.Filename = accFileName.Replace(".zip", ".txt");
                        newRecallFile.AccType = fileSplit[3];
                        newRecallFile.FileDate = fileSplit[4];
                        dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                //      newRecallFile.DateF = Convert.ToInt32(dateNew) + 2;
                if (noFields < 6)
                {
                    newRecallFile.Book = fileSplit[2];
                    newRecallFile.EDCWorkGroup = parts[0];
                    newRecallFile.Accountnumber = parts[1];
                    newRecallFile.LegalAccountNumber = parts[2];
                    newRecallFile.PrimaryPlasticNumber = parts[3];
                    newRecallFile.RecallDate = Convert.ToDateTime(parts[4]);
                }
                else
                {
                    newRecallFile.Book = fileSplit[2];
                    newRecallFile.EDCWorkGroup = parts[0];
                    newRecallFile.Accountnumber = parts[1];
                    newRecallFile.LegalAccountNumber = parts[2];
                    newRecallFile.PrimaryPlasticNumber = parts[3];
                    newRecallFile.RecallDate = Convert.ToDateTime(parts[4]);
                    newRecallFile.Recall_Area = parts[5];
                    newRecallFile.Prev_Risk = parts[6];
                    newRecallFile.Current_Risk = parts[7];


                }

            }
                
                DateTime dtInsertStarted = DateTime.Now;

                //  Insert this List<> of records using a Bulk Insert
             

                TimeSpan duration = DateTime.Now - dtInsertStarted;
                txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds","" , duration.Seconds, duration.Milliseconds) + Environment.NewLine;
            
            this.Cursor = Cursors.Arrow;
        }

        private void BtnAddList_Click(object sender, EventArgs e)
        {
            lstBxAccount.Items.Add(txtNewFile.Text);
            nList.Add(txtNewFile.Text);
        }

        private void btnClearList_Click(object sender, EventArgs e)
        {
            lstBxAccount.Items.Clear();
        }

        private void btnSingleProcess_Click(object sender, EventArgs e)
        {

            int i = 1;
            string[] fileSplit;
            string[] extensionsArray = { "_PAYMENT", "_RECALL", "_ACCOUNT", "_" };
            foreach (var itemAcc in nList)
            {
                var fn = itemAcc.ToString();
                if (extensionsArray.Any(fn.Contains))
                {
                    DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
                    txtAttach.Text += itemAcc.ToString() + Environment.NewLine;
                    if (File.Exists(basePath + itemAcc.ToString()))
                    {
                        ProcessFiles(itemAcc.ToString(), basePath);
                        fileSplit = itemAcc.Split('_');
                        var q = from p in BdSuiteContext.AbsaFileLogs
                                where p.Filename == itemAcc.Replace(".zip", ".txt")
                                select p;
                        if (q.Count() == 0)


                        {
                            AbsaFileLog objFilelog = new AbsaFileLog();
                            objFilelog.Filename = itemAcc;
                            objFilelog.AccType = "ACCOUNT";
                            objFilelog.FileDate = fileSplit[3];
                            objFilelog.Book = fileSplit[2];
                            objFilelog.CompPath = basePath;
                            objFilelog.ServPath = servPath;
                            objFilelog.Processed = 1;
                            BdSuiteContext.AbsaFileLogs.InsertOnSubmit(objFilelog);
                            BdSuiteContext.SubmitChanges();
                        }
                        else
                        {

                            var r = (from p in BdSuiteContext.AbsaFileLogs
                                     where p.Filename == itemAcc
                                     select p).SingleOrDefault();

                            r.Processed = 1;
                            BdSuiteContext.SubmitChanges();
                        }

                        if (i <= lstBxAccount.Items.Count - 1)
                        {
                            lstBxAccount.SelectedIndex = i;
                            i++;
                        }
                    }
                    else
                    {
                        var q = (from p in BdSuiteContext.AbsaFileLogs
                                 where p.Filename == itemAcc
                                 select p).SingleOrDefault();
                        q.Processed = 1;
                        BdSuiteContext.SubmitChanges();
                    }
                }
                break;
            }
            MessageBox.Show("Completed Single Processing File");

        }

        private void BtnMarkRedo_Click(object sender, EventArgs e)
        {
            String itemAcc;

            itemAcc = nList[0];

            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            
                            var r = (from p in BdSuiteContext.AbsaFileLogs
                                     where p.Filename == itemAcc
                                     select p).SingleOrDefault();

                            r.Processed = 3;
                            BdSuiteContext.SubmitChanges();
                        
            lstBxAccount.Items.Clear();
           nList.Clear();
            var q = from p in BdSuiteContext.AbsaFileLogs
                    where p.Processed == 0
                    orderby p.FileDate ascending, p.Book ascending
                    select p;
            foreach (var log in q)
            {

                lstBxAccount.Items.Add(log.Filename);
                nList.Add(log.Filename);

            }
            MessageBox.Show("Completed Mark Redo");

        }

        private void btnGetProcessRedo_Click(object sender, EventArgs e)
        {
            lstBxAccount.Items.Clear();
            nList.Clear();
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            var q = from p in BdSuiteContext.AbsaFileLogs
                    where p.Processed == 3
                    orderby p.FileDate ascending, p.Book ascending
                    select p;
            foreach (var log in q)
            {

                lstBxAccount.Items.Add(log.Filename);
                nList.Add(log.Filename);

            }
            MessageBox.Show("Completed Getting Unprocessed File");
        }

        private void btnGetoldUnprocess_Click(object sender, EventArgs e)
        {
            DateTime dt5 = dtePickNewDate.Value;
            DateTime dt3 = new DateTime(1989,12,31);

            // you can just convert it to integer if there is a constraint

            int dtI = (int)(dt5.ToOADate());

            lstBxAccountOld.Items.Clear();
            nListRedo.Clear();
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            var q = from p in BdSuiteContext.AbsaFileLogs
                    where p.Processed == 1 && p.Book == txtBoxBook.Text && p.AccType == txtBoxFileType.Text && p.DateF == dtI
                    orderby p.FileDate ascending, p.Book ascending
                    select p;
            foreach (var log in q)
            {

                lstBxAccountOld.Items.Add(log.Filename);
                nListRedo.Add(log.Filename);

            }
            MessageBox.Show("Completed Getting OLD Unprocessed File");
        }

        protected void CheckProcessFilesFix(String processFileName, String NewPath)

        {
            String[] fileSplit;

            fileSplit = processFileName.Split('_');
            if (fileSplit.Length > 2)
            {
                if (fileSplit[3] == "ACCOUNT")
                {
                    CheckProcessAccFileFix(processFileName, NewPath);
                }
                if (fileSplit[3] == "RECALL")
                {
                    CheckProcessRecallFileFix(processFileName, NewPath);
                }
                if (fileSplit[3] == "PAYMENT")
                {
                    CheckProcessPayFileFix(processFileName, NewPath);
                }
            }
        }


        private void btnFixProcess_Click(object sender, EventArgs e)
        {
            int i = 1;
            string[] extensionsArray = { "_PAYMENT", "_RECALL", "_ACCOUNT", "_" };
            foreach (var itemAcc in nListRedo)
            {
                var fn = itemAcc.ToString();
                if (extensionsArray.Any(fn.Contains))
                {
                    DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
                    txtAttach.Text += itemAcc.ToString() + Environment.NewLine;
                    if (File.Exists(basePath + itemAcc.ToString()))
                    {
                        CheckProcessFilesFix(itemAcc.ToString(), basePath);
                       

                        if (i <= lstBxAccountOld.Items.Count - 1)
                        {
                            lstBxAccountOld.SelectedIndex = i;
                            i++;
                        }
                    }

                }

            }
            MessageBox.Show("Completed Processing Fix File");
        }

        protected void CheckProcessAccFileFix(String accFileName, String NewPath)
        {
            String[] fileSplit;
            int noFields;
            int noUpd = 0;
            int allCount = 0;
            String newdata;

            DateTime dateNew;

            fileSplit = accFileName.Split('_');

            var textFieldParser = new TextFieldParser(NewPath + accFileName);

            textFieldParser.TextFieldType = FieldType.Delimited;
            textFieldParser.Delimiters = new[] { "|" };
            textFieldParser.TrimWhiteSpace = true;
            textFieldParser.HasFieldsEnclosedInQuotes = true;
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();

            // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
            while (!textFieldParser.EndOfData)
            {
                try
                {
                    noUpd++;
                    allCount++;
                    string[] parts = textFieldParser.ReadFields();
                    noFields = parts.Length;
                    txtProcess.Text = accFileName + " - " + noFields.ToString();
               

                    var newAccfile = (from p in BdSuiteContext.ABSAAccFiles
                                       where p.Filename == accFileName.Replace(".zip", ".txt") && p.Accountnumber == parts[1] 
                                      select p).SingleOrDefault();
                    if (newAccfile != null)
                    {

                        newAccfile.Filename = accFileName.Replace(".zip", ".txt");
                        newAccfile.AccType = fileSplit[3];
                        newAccfile.FileDate = fileSplit[4];
                        dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        //    newAccfile.DateF = Convert.tol(dateNew) + 2;
                        newAccfile.Book = fileSplit[2];
                        newAccfile.Ind = parts[0];
                        newAccfile.Accountnumber = parts[1];
                        newAccfile.Title = parts[2];
                        newAccfile.Initials = parts[3];
                        newAccfile.Surname = parts[4];
                        newAccfile.IDNumber = parts[5];
                        newAccfile.EmailAddress = parts[6];
                        newAccfile.CellNo = parts[7];
                        newAccfile.HomeTel = parts[8];
                        newAccfile.WorkTel = parts[9];
                        newAccfile.FaxNo = parts[10];
                        newAccfile.HomeLine1 = parts[11];
                        newAccfile.HomeLine2 = parts[12];
                        newAccfile.HomeLine3 = parts[13];
                        newAccfile.HomeLine4 = parts[14];
                        newAccfile.HomePostal = parts[15];
                        newAccfile.PostalLine1 = parts[16];
                        newAccfile.PostalLine2 = parts[17];
                        newAccfile.PostalLine3 = parts[18];
                        newAccfile.PostalLine4 = parts[19];
                        newAccfile.PostalCode = parts[20];
                        newAccfile.EmployerName = parts[21];
                        newAccfile.EmployerAddLine1 = parts[22];
                        newAccfile.EmployerAddLine2 = parts[23];
                        newAccfile.EmployerAddLine3 = parts[24];
                        newAccfile.EmployerPostalCode = parts[25];
                        newAccfile.Occupation = parts[26];
                        newAccfile.Product = parts[27];
                        newAccfile.InstructionReceivedDate = parts[28];
                        newAccfile.InterestFromDate = parts[29];
                        newdata = parts[30];
                        if (!string.IsNullOrEmpty(parts[30]) || parts[30] == "0.00")
                        {
                            newAccfile.InterestRate = Convert.ToDecimal(parts[30]);
                        }
                        newAccfile.SBUCode = parts[31];
                        newAccfile.BranchNr = parts[32];
                        newAccfile.LastPaymentDate = parts[33];
                        if (!string.IsNullOrEmpty(parts[34]) || parts[34] == "0.00")
                        {

                            newAccfile.LastPaymentAmount = Convert.ToDecimal(parts[34]);
                        }
                        newAccfile.Tag1 = parts[35];
                        newAccfile.Tagassignmentdate1 = parts[36];
                        newAccfile.Tag2 = parts[37];
                        newAccfile.Tagassignmentdate2 = parts[38];
                        newAccfile.Tag3 = parts[39];
                        newAccfile.Tagassignmentdate3 = parts[40];
                        newAccfile.Tag4 = parts[41];
                        newAccfile.Tagassignmentdate4 = parts[42];
                        newAccfile.Tag5 = parts[43];
                        newAccfile.Tagassignmentdate5 = parts[44];
                        newAccfile.SummonsCourt = parts[45];
                        newAccfile.SummonsCaseNumber = parts[46];
                        newAccfile.SummonsServedDate = parts[47];
                        newAccfile.JudgementCaseNumber = parts[48];
                        newAccfile.JudgementCourtDate = parts[49];
                        newAccfile.LegalJudgementDAte = parts[50];
                        if (!string.IsNullOrEmpty(parts[51]))
                        {
                            newAccfile.DefaultJudgementAmount = Convert.ToDecimal(parts[51]);
                        }
                        newAccfile.SIEOUTCOMEDateSold = parts[52];
                        if (!string.IsNullOrEmpty(parts[53]))
                        {

                            newAccfile.SIEOUTCOMESoldAmount = Convert.ToDecimal(parts[53]);
                        }
                        newAccfile.DateWrittenOff = parts[54];
                        if (!string.IsNullOrEmpty(parts[55]))
                        {
                            newAccfile.CapitalWrittenOff = Convert.ToDecimal(parts[55]);
                        }
                        if (!string.IsNullOrEmpty(parts[56]))
                        {
                            newAccfile.InterestWrittenOff = Convert.ToDecimal(parts[56]);
                        }
                        newAccfile.RecoverySummonsCourt = parts[57];
                        newAccfile.RecoverySummonsCase = parts[58];
                        newAccfile.RecoverySummonsServedDate = parts[59];
                        newAccfile.RecoveryJudgementDate = parts[60];
                        newAccfile.RecoverySIEOUTCOMEDateSold = parts[61];
                        if (!string.IsNullOrEmpty(parts[62]))
                        {
                            newAccfile.RecoverySIEOUTCOMESoldAmount = Convert.ToDecimal(parts[62]);
                        }
                        if (!string.IsNullOrEmpty(parts[63]))
                        {
                            newAccfile.NettCredits = Convert.ToDecimal(parts[63]);
                        }
                        if (!string.IsNullOrEmpty(parts[64]))
                        {

                            newAccfile.Arrears = Convert.ToDecimal(parts[64]);
                        }



                        if (!string.IsNullOrEmpty(parts[65]))
                        {
                            newAccfile.Balance = Convert.ToDecimal(parts[65]);
                        }
                        newAccfile.Asset = parts[66];
                        if (noFields > 67)
                        {
                            newAccfile.PrelegalInstructionReceivedDate = parts[67];
                            newAccfile.PrelegalName = parts[68];
                            newAccfile.LegalInstructionReceivedDate = parts[69];
                            newAccfile.LegalName = parts[70];
                        }
                        if (noFields > 71)
                        {


                            if (!string.IsNullOrEmpty(parts[71]))
                            {
                                newAccfile.InstalmentAmount = Convert.ToDecimal(parts[71]);
                            }
                            newAccfile.InstalmentDueDate = parts[72];
                            newAccfile.InstalmentDate = parts[73];
                            newAccfile.DateaccountOpened = parts[74];
                            newAccfile.NID = parts[75];
                            newAccfile.NDD = parts[76];
                            newAccfile.MOB = parts[77];
                            if (!string.IsNullOrEmpty(parts[78]))
                            {
                                newAccfile.OriginalBalance = Convert.ToDecimal(parts[78]);
                            }
                            if (!string.IsNullOrEmpty(parts[79]))
                            {

                                newAccfile.OutstandingBalance = Convert.ToDecimal(parts[79]);
                            }
                            newAccfile.Tag6 = parts[80];
                            newAccfile.Tag6AssignmentDate = parts[81];
                            newAccfile.Tag7 = parts[82];
                            newAccfile.Tag7AssignmentDate = parts[83];
                            newAccfile.Tag8 = parts[84];
                            newAccfile.Tag8AssignmentDate = parts[85];
                            if (!string.IsNullOrEmpty(parts[86]))
                            {
                                newAccfile.IntRate = Convert.ToDecimal(parts[86]);
                            }
                            newAccfile.LoanTerm = parts[87];
                            newAccfile.NrPTPCCC = parts[88];
                            newAccfile.NrPTPConsec = parts[89];
                            newAccfile.NrPTPLife = parts[90];
                            newAccfile.BranchName = parts[91];
                        }
                        if (noFields > 92)
                        {

                            newAccfile.RiskType = parts[92];
                            newAccfile.TRIADPool = parts[93];
                            newAccfile.CreditLifeapplicable = parts[94];
                            newAccfile.PlasticNr = parts[95];
                            newAccfile.LegalAccountNBumber = parts[96];
                            newAccfile.Unemployed = parts[97];
                            newAccfile.Unemployedsince = parts[98];
                            newAccfile.UnemploymentAffidavitDate = parts[99];
                            newAccfile.Section129Letter = parts[100];
                            newAccfile.varchargeoffDate = parts[101];
                            if (!string.IsNullOrEmpty(parts[102]))
                            {
                                newAccfile.ArrearBalanceatvarchargeoff = Convert.ToDecimal(parts[102]);
                            }
                            if (!string.IsNullOrEmpty(parts[103]))
                            {
                                newAccfile.OutstandingBalanceatvarchargeoff = Convert.ToDecimal(parts[103]);
                            }
                            if (!string.IsNullOrEmpty(parts[104]))
                            {
                                newAccfile.Interestatvarchargeoff = Convert.ToDecimal(parts[104]);
                            }
                            if (!string.IsNullOrEmpty(parts[105]))
                            {

                                newAccfile.HandoverAmount = Convert.ToDecimal(parts[105]);
                            }
                            newAccfile.AssetDescription = parts[106];
                            newAccfile.AssetRegistration = parts[107];
                            newAccfile.Enginenumber = parts[108];
                            newAccfile.Chassisnumber = parts[109];
                            newAccfile.Vinnumber = parts[110];
                            newAccfile.MMCode = parts[111];
                            newAccfile.DealerName = parts[112];
                            newAccfile.DealerNumber = parts[113];
                            if (!string.IsNullOrEmpty(parts[114]))
                            {
                                newAccfile.TradeValue = Convert.ToDecimal(parts[114]);
                            }
                            if (!string.IsNullOrEmpty(parts[115]))
                            {

                                newAccfile.RetailValue = Convert.ToDecimal(parts[115]);
                            }
                            if (!string.IsNullOrEmpty(parts[116]))
                            {
                                newAccfile.SoldAmount = Convert.ToDecimal(parts[116]);
                            }
                            newAccfile.TradeCentreNumber = parts[117];
                            newAccfile.VehicleStockNumber = parts[118];
                            newAccfile.Assetinpossession = parts[119];
                            newAccfile.WarrantIssued = parts[120];
                            newAccfile.Warrant_Sheriff = parts[121].Trim();
                            newAccfile.Sheriffappointed = parts[122].Trim();
                            newAccfile.DocumentationCompleted = parts[123].Trim();
                            newAccfile.DisputeStatus = parts[124];
                            newAccfile.DisputeStatusDate = parts[125];
                            newAccfile.Complaint = parts[126];
                            newAccfile.Complaintdate = parts[127];
                            newAccfile.DebtReviewIndicator = parts[128];
                            newAccfile.InternalPlacementbefore = parts[129];
                            newAccfile.InternalPlacementdate = parts[130];
                            newAccfile.Prescriptionaccount = parts[131];
                            newAccfile.LastPTP = parts[132];
                        }
                        if (noFields > 133)
                        {
                            newAccfile.FirstName = parts[133];
                            if (!string.IsNullOrEmpty(parts[134]))
                            {
                                newAccfile.SettlmentPercentage = Convert.ToDecimal(parts[134]);
                            }
                            if (!string.IsNullOrEmpty(parts[135]))
                            {
                                newAccfile.PastDue = Convert.ToDecimal(parts[135]);
                            }
                            if (!string.IsNullOrEmpty(parts[136]))
                            {
                                newAccfile.Limit = Convert.ToDecimal(parts[136]);
                            }
                            if (!string.IsNullOrEmpty(parts[137]))
                            {
                                newAccfile.OverLimit = Convert.ToDecimal(parts[137]);
                            }
                            if (!string.IsNullOrEmpty(parts[138]))
                            {
                                newAccfile.VAtAmt = Convert.ToDecimal(parts[138]);
                            }
                        }
                        if (noFields > 139)
                        {

                            if (!string.IsNullOrEmpty(parts[139]))
                            {
                                newAccfile.ECBalance = Convert.ToDecimal(parts[139]);
                            }
                            if (!string.IsNullOrEmpty(parts[140]))
                            {
                                newAccfile.CONSBalance = Convert.ToDecimal(parts[140]);
                            }
                            if (!string.IsNullOrEmpty(parts[141]))
                            {
                                newAccfile.PrePrescribeTag = parts[141];
                            }
                        }
                        if (noFields > 142)
                        {
                            if (!string.IsNullOrEmpty(parts[142]))
                            {
                                newAccfile.Total_OutStanding_Bal = Convert.ToDecimal(parts[142]);
                            }
                            if (!string.IsNullOrEmpty(parts[143]))
                            {
                                newAccfile.OverDarft_Interest_Rate = Convert.ToDecimal(parts[143]);
                            }
                            if (!string.IsNullOrEmpty(parts[144]))
                            {
                                newAccfile.Prev_Install_Amount = Convert.ToDecimal(parts[143]);
                            }
                            if (!string.IsNullOrEmpty(parts[145]))
                            {
                                newAccfile.Decline_Reason = (parts[145]);
                            }
                            if (!string.IsNullOrEmpty(parts[146]))
                            {
                                newAccfile.Memo = (parts[146]);
                            }
                            if (!string.IsNullOrEmpty(parts[147]))
                            {
                                newAccfile.XDS_ITC1 = parts[147];
                            }

                            if (!string.IsNullOrEmpty(parts[148]))
                            {
                                newAccfile.XDS_ITC2 = (parts[148]);
                            }
                            if (!string.IsNullOrEmpty(parts[149]))
                            {
                                newAccfile.EDC_GIVEN1 = (parts[149]);
                            }
                            if (!string.IsNullOrEmpty(parts[150]))
                            {
                                newAccfile.EDC_GIVEN2 = parts[150];
                            }
                            if (!string.IsNullOrEmpty(parts[151]))
                            {
                                newAccfile.INTERNAL_PHN1 = (parts[151]);
                            }
                            if (!string.IsNullOrEmpty(parts[152]))
                            {
                                newAccfile.INTERNAL_PHN2 = (parts[152]);
                            }
                            if (!string.IsNullOrEmpty(parts[153]))
                            {
                                newAccfile.INTERNAL_PHN3 = parts[153];
                            }
                            if (!string.IsNullOrEmpty(parts[154]))
                            {
                                newAccfile.INTERNAL_PHN4 = (parts[154]);
                            }
                            if (!string.IsNullOrEmpty(parts[155]))
                            {
                                newAccfile.INTERNAL_PHN5 = (parts[155]);
                            }
                            if (!string.IsNullOrEmpty(parts[156]))
                            {
                                newAccfile.INTERNAL_PHN6 = parts[156];
                            }


                        }
                    
                        BdSuiteContext.SubmitChanges();
                    }
                }
                catch (System.FormatException ex3)
                {
                    txtAttach.Text += ex3.Message;
                    txtAttach.Text += string.Format("Error record not imported format error ") + Environment.NewLine;
                }
            }



            DateTime dtInsertStarted = DateTime.Now;
            TimeSpan duration = DateTime.Now - dtInsertStarted;
            txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", "", duration.Seconds, duration.Milliseconds) + Environment.NewLine;

            this.Cursor = Cursors.Arrow;
        }


        protected void CheckProcessPayFileFix(String accFileName, String NewPath)
        {
            String[] fileSplit;
            int noFields;
            CultureInfo provider = CultureInfo.InvariantCulture;
            DateTime dateNew;

            int lineno = 0;

            fileSplit = accFileName.Split('_');
            String checkStr;
            String checkData;

            var textFieldParser = new TextFieldParser(NewPath + accFileName);

            textFieldParser.TextFieldType = FieldType.Delimited;
            textFieldParser.Delimiters = new[] { "|" };
            textFieldParser.TrimWhiteSpace = true;
            textFieldParser.HasFieldsEnclosedInQuotes = true;
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
            while (!textFieldParser.EndOfData)
            {

                lineno++;
                string[] parts = textFieldParser.ReadFields();
                noFields = parts.Length;

                txtProcess.Text = accFileName + " - " + noFields.ToString();

                var newPayFile = (from p in BdSuiteContext.ABSAPayFiles
                                  where p.Filename == accFileName.Replace(".zip", ".txt") && p.AccountNumber == parts[0] && p.Amount == Convert.ToDecimal(parts[1])
                                  && p.FinancialTransactionEffectiveDate == Convert.ToDateTime(parts[3]) && p.PaymentType == parts[5]
                                  select p).SingleOrDefault();

                newPayFile.Filename = accFileName.Replace(".zip", ".txt");
                newPayFile.AccType = fileSplit[3];
                dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                //   newPayFile.DateF = Convert.ToInt32(dateNew) + 2;
                newPayFile.FileDate = fileSplit[4];
                newPayFile.Book = fileSplit[2];

                checkData = parts[2];
                checkStr = checkData;

                if (noFields < 19)
                {

                    newPayFile.AccountNumber = parts[0];
                    newPayFile.Amount = Convert.ToDecimal(parts[1]);
                    if (!string.IsNullOrEmpty(parts[2]) && parts[2].Length > 9)
                    {
                        //  newPayFile.FinancialTransactionEffectiveDate = DateTime.ParseExact(parts[2], "yyyy/MM/dd", provider);
                    }
                    newPayFile.MemoCode = parts[3];
                    newPayFile.PaymentType = parts[4];
                    newPayFile.PaymentMethod = parts[5];
                }
                else
                {


                    newPayFile.Workgroup = parts[0];
                    newPayFile.AccountNumber = parts[1];
                    newPayFile.Amount = Convert.ToDecimal(parts[2]);
                    if (!string.IsNullOrEmpty(parts[3]))
                    {
                        newPayFile.FinancialTransactionEffectiveDate = Convert.ToDateTime(parts[3]);
                    }
                    newPayFile.MemoCode = parts[4];
                    newPayFile.PaymentType = parts[5];
                    newPayFile.PaymentMethod = parts[6];
                    // Tag6 = parts[7],
                    //  Tag6AssignmentDate = Convert.ToDateTime(parts[8]),
                    newPayFile.Tag4 = parts[9];
                    if (!string.IsNullOrEmpty(parts[10]))
                    {
                        newPayFile.Tag4AssignmentDate = Convert.ToDateTime(parts[10]);
                    }
                    newPayFile.Tag5 = parts[11];
                    if (!string.IsNullOrEmpty(parts[12]))
                    {
                        newPayFile.Tag5AssignmentDate = Convert.ToDateTime(parts[12]);
                    }
                    newPayFile.NAEDO = parts[13];
                    if (!string.IsNullOrEmpty(parts[14]))
                    {
                        newPayFile.Dateaccountopened = Convert.ToDateTime(parts[14]);
                    }
                    newPayFile.SiteName = parts[15];
                    newPayFile.IDNumber = parts[16];
                    newPayFile.Surname = parts[17];
                    newPayFile.SBUCode = parts[18];
                    newPayFile.Site = parts[19];
                }
                if (noFields >= 21)
                {
                    newPayFile.AttorneyCode = parts[20];
                    newPayFile.Commissionable = parts[21];
                }
                BdSuiteContext.SubmitChanges();

            }



            DateTime dtInsertStarted = DateTime.Now;

            TimeSpan duration = DateTime.Now - dtInsertStarted;
            txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", "", duration.Seconds, duration.Milliseconds) + Environment.NewLine;

            this.Cursor = Cursors.Arrow;
        }

        protected void CheckProcessRecallFileFix(String accFileName, String NewPath)
        {
            String[] fileSplit;
            CultureInfo provider = CultureInfo.InvariantCulture;
            DateTime dateNew;
            fileSplit = accFileName.Split('_');
            int noFields;
            var textFieldParser = new TextFieldParser(NewPath + accFileName);

            textFieldParser.TextFieldType = FieldType.Delimited;
            textFieldParser.Delimiters = new[] { "|" };
            textFieldParser.TrimWhiteSpace = true;
            textFieldParser.HasFieldsEnclosedInQuotes = true;
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
            while (!textFieldParser.EndOfData)
            {
                string[] parts = textFieldParser.ReadFields();
                noFields = parts.Length;
              
                var newRecallFile = (from p in BdSuiteContext.ABSARecallFiles
                                     where p.Filename == accFileName.Replace(".zip", ".txt") && p.Accountnumber == parts[1]
                                     select p).SingleOrDefault();

                newRecallFile.Filename = accFileName.Replace(".zip", ".txt");
                newRecallFile.AccType = fileSplit[3];
                newRecallFile.FileDate = fileSplit[4];
                dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                //      newRecallFile.DateF = Convert.ToInt32(dateNew) + 2;
                if (noFields < 6)
                {
                    newRecallFile.Book = fileSplit[2];
                    newRecallFile.EDCWorkGroup = parts[0];
                    newRecallFile.Accountnumber = parts[1];
                    newRecallFile.LegalAccountNumber = parts[2];
                    newRecallFile.PrimaryPlasticNumber = parts[3];
                    newRecallFile.RecallDate = Convert.ToDateTime(parts[4]);
                }
                else
                {
                    newRecallFile.Book = fileSplit[2];
                    newRecallFile.EDCWorkGroup = parts[0];
                    newRecallFile.Accountnumber = parts[1];
                    newRecallFile.LegalAccountNumber = parts[2];
                    newRecallFile.PrimaryPlasticNumber = parts[3];
                    newRecallFile.RecallDate = Convert.ToDateTime(parts[4]);
                    newRecallFile.Recall_Area = parts[5];
                    newRecallFile.Prev_Risk = parts[6];
                    newRecallFile.Current_Risk = parts[7];


                }
                BdSuiteContext.SubmitChanges();

            }

            DateTime dtInsertStarted = DateTime.Now;

            //  Insert this List<> of records using a Bulk Insert


            TimeSpan duration = DateTime.Now - dtInsertStarted;
            txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", "", duration.Seconds, duration.Milliseconds) + Environment.NewLine;

            this.Cursor = Cursors.Arrow;
        }

        protected void CheckProcessFilesFixDelRedo(String processFileName, String NewPath)

        {
            String[] fileSplit;

            fileSplit = processFileName.Split('_');
            if (fileSplit.Length > 2)
            {
                if (fileSplit[3] == "ACCOUNT")
                {
                    CheckProcessAccFileRedo(processFileName, NewPath);
                }
                if (fileSplit[3] == "RECALL")
                {
                    CheckProcessRecallFileRedo(processFileName, NewPath);
                }
                if (fileSplit[3] == "PAYMENT")
                {
                    CheckProcessPayFileRedo(processFileName, NewPath);
                }
            }
        }


        private void bntoldProcess_Click(object sender, EventArgs e)
        {
            int i = 1;
           
            string[] extensionsArray = { "_PAYMENT", "_RECALL", "_ACCOUNT", "_" };
            foreach (var itemAcc in nListRedo)
            {
                var fn = itemAcc.ToString();
                if (extensionsArray.Any(fn.Contains))
                {
                    DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
                    txtAttach.Text += itemAcc.ToString() + Environment.NewLine;
                    if (File.Exists(basePath + itemAcc.ToString()))
                    {
                        CheckProcessFilesFixDelRedo(itemAcc.ToString(), basePath);
                      
                        if (i <= lstBxAccountOld.Items.Count - 1)
                        {
                            lstBxAccountOld.SelectedIndex = i;
                            i++;
                        }
                    }
 
                }

            }
            MessageBox.Show("Completed Processing Redo File");
        }

        protected void CheckProcessAccFileRedo(String accFileName, String NewPath)
        {
            String[] fileSplit;
            int noFields;
            int noUpd = 0;
            int allCount = 0;
            String newdata;

            DateTime dateNew;

            fileSplit = accFileName.Split('_');

            var textFieldParser = new TextFieldParser(NewPath + accFileName);

            textFieldParser.TextFieldType = FieldType.Delimited;
            textFieldParser.Delimiters = new[] { "|" };
            textFieldParser.TrimWhiteSpace = true;
            textFieldParser.HasFieldsEnclosedInQuotes = true;
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
            while (!textFieldParser.EndOfData)
            {
                try
                {
                    noUpd++;
                    allCount++;
                    string[] parts = textFieldParser.ReadFields();
                    noFields = parts.Length;
                    txtProcess.Text = accFileName + " - " + noFields.ToString();
                    var newAccfile = (from p in BdSuiteContext.ABSAAccFiles
                                      where p.Filename == accFileName.Replace(".zip", ".txt") && p.Accountnumber == parts[1]
                                      select p).SingleOrDefault();


                    newAccfile.Filename = accFileName.Replace(".zip", ".txt");
                    newAccfile.AccType = fileSplit[3];
                    newAccfile.FileDate = fileSplit[4];
                    dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                    //    newAccfile.DateF = Convert.tol(dateNew) + 2;
                    newAccfile.Book = fileSplit[2];
                    newAccfile.Ind = parts[0];
                    newAccfile.Accountnumber = parts[1];
                    newAccfile.Title = parts[2];
                    newAccfile.Initials = parts[3];
                    newAccfile.Surname = parts[4];
                    newAccfile.IDNumber = parts[5];
                    newAccfile.EmailAddress = parts[6];
                    newAccfile.CellNo = parts[7];
                    newAccfile.HomeTel = parts[8];
                    newAccfile.WorkTel = parts[9];
                    newAccfile.FaxNo = parts[10];
                    newAccfile.HomeLine1 = parts[11];
                    newAccfile.HomeLine2 = parts[12];
                    newAccfile.HomeLine3 = parts[13];
                    newAccfile.HomeLine4 = parts[14];
                    newAccfile.HomePostal = parts[15];
                    newAccfile.PostalLine1 = parts[16];
                    newAccfile.PostalLine2 = parts[17];
                    newAccfile.PostalLine3 = parts[18];
                    newAccfile.PostalLine4 = parts[19];
                    newAccfile.PostalCode = parts[20];
                    newAccfile.EmployerName = parts[21];
                    newAccfile.EmployerAddLine1 = parts[22];
                    newAccfile.EmployerAddLine2 = parts[23];
                    newAccfile.EmployerAddLine3 = parts[24];
                    newAccfile.EmployerPostalCode = parts[25];
                    newAccfile.Occupation = parts[26];
                    newAccfile.Product = parts[27];
                    newAccfile.InstructionReceivedDate = parts[28];
                    newAccfile.InterestFromDate = parts[29];
                    newdata = parts[30];
                    if (!string.IsNullOrEmpty(parts[30]) || parts[30] == "0.00")
                    {
                        newAccfile.InterestRate = Convert.ToDecimal(parts[30]);
                    }
                    newAccfile.SBUCode = parts[31];
                    newAccfile.BranchNr = parts[32];
                    newAccfile.LastPaymentDate = parts[33];
                    if (!string.IsNullOrEmpty(parts[34]) || parts[34] == "0.00")
                    {

                        newAccfile.LastPaymentAmount = Convert.ToDecimal(parts[34]);
                    }
                    newAccfile.Tag1 = parts[35];
                    newAccfile.Tagassignmentdate1 = parts[36];
                    newAccfile.Tag2 = parts[37];
                    newAccfile.Tagassignmentdate2 = parts[38];
                    newAccfile.Tag3 = parts[39];
                    newAccfile.Tagassignmentdate3 = parts[40];
                    newAccfile.Tag4 = parts[41];
                    newAccfile.Tagassignmentdate4 = parts[42];
                    newAccfile.Tag5 = parts[43];
                    newAccfile.Tagassignmentdate5 = parts[44];
                    newAccfile.SummonsCourt = parts[45];
                    newAccfile.SummonsCaseNumber = parts[46];
                    newAccfile.SummonsServedDate = parts[47];
                    newAccfile.JudgementCaseNumber = parts[48];
                    newAccfile.JudgementCourtDate = parts[49];
                    newAccfile.LegalJudgementDAte = parts[50];
                    if (!string.IsNullOrEmpty(parts[51]))
                    {
                        newAccfile.DefaultJudgementAmount = Convert.ToDecimal(parts[51]);
                    }
                    newAccfile.SIEOUTCOMEDateSold = parts[52];
                    if (!string.IsNullOrEmpty(parts[53]))
                    {

                        newAccfile.SIEOUTCOMESoldAmount = Convert.ToDecimal(parts[53]);
                    }
                    newAccfile.DateWrittenOff = parts[54];
                    if (!string.IsNullOrEmpty(parts[55]))
                    {
                        newAccfile.CapitalWrittenOff = Convert.ToDecimal(parts[55]);
                    }
                    if (!string.IsNullOrEmpty(parts[56]))
                    {
                        newAccfile.InterestWrittenOff = Convert.ToDecimal(parts[56]);
                    }
                    newAccfile.RecoverySummonsCourt = parts[57];
                    newAccfile.RecoverySummonsCase = parts[58];
                    newAccfile.RecoverySummonsServedDate = parts[59];
                    newAccfile.RecoveryJudgementDate = parts[60];
                    newAccfile.RecoverySIEOUTCOMEDateSold = parts[61];
                    if (!string.IsNullOrEmpty(parts[62]))
                    {
                        newAccfile.RecoverySIEOUTCOMESoldAmount = Convert.ToDecimal(parts[62]);
                    }
                    if (!string.IsNullOrEmpty(parts[63]))
                    {
                        newAccfile.NettCredits = Convert.ToDecimal(parts[63]);
                    }
                    if (!string.IsNullOrEmpty(parts[64]))
                    {

                        newAccfile.Arrears = Convert.ToDecimal(parts[64]);
                    }



                    if (!string.IsNullOrEmpty(parts[65]))
                    {
                        newAccfile.Balance = Convert.ToDecimal(parts[65]);
                    }
                    newAccfile.Asset = parts[66];
                    if (noFields > 67)
                    {
                        newAccfile.PrelegalInstructionReceivedDate = parts[67];
                        newAccfile.PrelegalName = parts[68];
                        newAccfile.LegalInstructionReceivedDate = parts[69];
                        newAccfile.LegalName = parts[70];
                    }
                    if (noFields > 71)
                    {


                        if (!string.IsNullOrEmpty(parts[71]))
                        {
                            newAccfile.InstalmentAmount = Convert.ToDecimal(parts[71]);
                        }
                        newAccfile.InstalmentDueDate = parts[72];
                        newAccfile.InstalmentDate = parts[73];
                        newAccfile.DateaccountOpened = parts[74];
                        newAccfile.NID = parts[75];
                        newAccfile.NDD = parts[76];
                        newAccfile.MOB = parts[77];
                        if (!string.IsNullOrEmpty(parts[78]))
                        {
                            newAccfile.OriginalBalance = Convert.ToDecimal(parts[78]);
                        }
                        if (!string.IsNullOrEmpty(parts[79]))
                        {

                            newAccfile.OutstandingBalance = Convert.ToDecimal(parts[79]);
                        }
                        newAccfile.Tag6 = parts[80];
                        newAccfile.Tag6AssignmentDate = parts[81];
                        newAccfile.Tag7 = parts[82];
                        newAccfile.Tag7AssignmentDate = parts[83];
                        newAccfile.Tag8 = parts[84];
                        newAccfile.Tag8AssignmentDate = parts[85];
                        if (!string.IsNullOrEmpty(parts[86]))
                        {
                            newAccfile.IntRate = Convert.ToDecimal(parts[86]);
                        }
                        newAccfile.LoanTerm = parts[87];
                        newAccfile.NrPTPCCC = parts[88];
                        newAccfile.NrPTPConsec = parts[89];
                        newAccfile.NrPTPLife = parts[90];
                        newAccfile.BranchName = parts[91];
                    }
                    if (noFields > 92)
                    {

                        newAccfile.RiskType = parts[92];
                        newAccfile.TRIADPool = parts[93];
                        newAccfile.CreditLifeapplicable = parts[94];
                        newAccfile.PlasticNr = parts[95];
                        newAccfile.LegalAccountNBumber = parts[96];
                        newAccfile.Unemployed = parts[97];
                        newAccfile.Unemployedsince = parts[98];
                        newAccfile.UnemploymentAffidavitDate = parts[99];
                        newAccfile.Section129Letter = parts[100];
                        newAccfile.varchargeoffDate = parts[101];
                        if (!string.IsNullOrEmpty(parts[102]))
                        {
                            newAccfile.ArrearBalanceatvarchargeoff = Convert.ToDecimal(parts[102]);
                        }
                        if (!string.IsNullOrEmpty(parts[103]))
                        {
                            newAccfile.OutstandingBalanceatvarchargeoff = Convert.ToDecimal(parts[103]);
                        }
                        if (!string.IsNullOrEmpty(parts[104]))
                        {
                            newAccfile.Interestatvarchargeoff = Convert.ToDecimal(parts[104]);
                        }
                        if (!string.IsNullOrEmpty(parts[105]))
                        {

                            newAccfile.HandoverAmount = Convert.ToDecimal(parts[105]);
                        }
                        newAccfile.AssetDescription = parts[106];
                        newAccfile.AssetRegistration = parts[107];
                        newAccfile.Enginenumber = parts[108];
                        newAccfile.Chassisnumber = parts[109];
                        newAccfile.Vinnumber = parts[110];
                        newAccfile.MMCode = parts[111];
                        newAccfile.DealerName = parts[112];
                        newAccfile.DealerNumber = parts[113];
                        if (!string.IsNullOrEmpty(parts[114]))
                        {
                            newAccfile.TradeValue = Convert.ToDecimal(parts[114]);
                        }
                        if (!string.IsNullOrEmpty(parts[115]))
                        {

                            newAccfile.RetailValue = Convert.ToDecimal(parts[115]);
                        }
                        if (!string.IsNullOrEmpty(parts[116]))
                        {
                            newAccfile.SoldAmount = Convert.ToDecimal(parts[116]);
                        }
                        newAccfile.TradeCentreNumber = parts[117];
                        newAccfile.VehicleStockNumber = parts[118];
                        newAccfile.Assetinpossession = parts[119];
                        newAccfile.WarrantIssued = parts[120];
                        newAccfile.Warrant_Sheriff = parts[121].Trim();
                        newAccfile.Sheriffappointed = parts[122].Trim();
                        newAccfile.DocumentationCompleted = parts[123].Trim();
                        newAccfile.DisputeStatus = parts[124];
                        newAccfile.DisputeStatusDate = parts[125];
                        newAccfile.Complaint = parts[126];
                        newAccfile.Complaintdate = parts[127];
                        newAccfile.DebtReviewIndicator = parts[128];
                        newAccfile.InternalPlacementbefore = parts[129];
                        newAccfile.InternalPlacementdate = parts[130];
                        newAccfile.Prescriptionaccount = parts[131];
                        newAccfile.LastPTP = parts[132];
                    }
                    if (noFields > 133)
                    {
                        newAccfile.FirstName = parts[133];
                        if (!string.IsNullOrEmpty(parts[134]))
                        {
                            newAccfile.SettlmentPercentage = Convert.ToDecimal(parts[134]);
                        }
                        if (!string.IsNullOrEmpty(parts[135]))
                        {
                            newAccfile.PastDue = Convert.ToDecimal(parts[135]);
                        }
                        if (!string.IsNullOrEmpty(parts[136]))
                        {
                            newAccfile.Limit = Convert.ToDecimal(parts[136]);
                        }
                        if (!string.IsNullOrEmpty(parts[137]))
                        {
                            newAccfile.OverLimit = Convert.ToDecimal(parts[137]);
                        }
                        if (!string.IsNullOrEmpty(parts[138]))
                        {
                            newAccfile.VAtAmt = Convert.ToDecimal(parts[138]);
                        }
                    }
                    if (noFields > 139)
                    {

                        if (!string.IsNullOrEmpty(parts[139]))
                        {
                            newAccfile.ECBalance = Convert.ToDecimal(parts[139]);
                        }
                        if (!string.IsNullOrEmpty(parts[140]))
                        {
                            newAccfile.CONSBalance = Convert.ToDecimal(parts[140]);
                        }
                        if (!string.IsNullOrEmpty(parts[141]))
                        {
                            newAccfile.PrePrescribeTag = parts[141];
                        }
                    }
                    if (noFields > 142)
                    {
                        if (!string.IsNullOrEmpty(parts[142]))
                        {
                            newAccfile.Total_OutStanding_Bal = Convert.ToDecimal(parts[142]);
                        }
                        if (!string.IsNullOrEmpty(parts[143]))
                        {
                            newAccfile.OverDarft_Interest_Rate = Convert.ToDecimal(parts[143]);
                        }
                        if (!string.IsNullOrEmpty(parts[144]))
                        {
                            newAccfile.Prev_Install_Amount = Convert.ToDecimal(parts[143]);
                        }
                        if (!string.IsNullOrEmpty(parts[145]))
                        {
                            newAccfile.Decline_Reason = (parts[145]);
                        }
                        if (!string.IsNullOrEmpty(parts[146]))
                        {
                            newAccfile.Memo = (parts[146]);
                        }
                        if (!string.IsNullOrEmpty(parts[147]))
                        {
                            newAccfile.XDS_ITC1 = parts[147];
                        }

                        if (!string.IsNullOrEmpty(parts[148]))
                        {
                            newAccfile.XDS_ITC2 = (parts[148]);
                        }
                        if (!string.IsNullOrEmpty(parts[149]))
                        {
                            newAccfile.EDC_GIVEN1 = (parts[149]);
                        }
                        if (!string.IsNullOrEmpty(parts[150]))
                        {
                            newAccfile.EDC_GIVEN2 = parts[150];
                        }
                        if (!string.IsNullOrEmpty(parts[151]))
                        {
                            newAccfile.INTERNAL_PHN1 = (parts[151]);
                        }
                        if (!string.IsNullOrEmpty(parts[152]))
                        {
                            newAccfile.INTERNAL_PHN2 = (parts[152]);
                        }
                        if (!string.IsNullOrEmpty(parts[153]))
                        {
                            newAccfile.INTERNAL_PHN3 = parts[153];
                        }
                        if (!string.IsNullOrEmpty(parts[154]))
                        {
                            newAccfile.INTERNAL_PHN4 = (parts[154]);
                        }
                        if (!string.IsNullOrEmpty(parts[155]))
                        {
                            newAccfile.INTERNAL_PHN5 = (parts[155]);
                        }
                        if (!string.IsNullOrEmpty(parts[156]))
                        {
                            newAccfile.INTERNAL_PHN6 = parts[156];
                        }


                    }
                    BdSuiteContext.SubmitChanges();

                }
                catch (System.FormatException ex3)
                {
                    txtAttach.Text += ex3.Message;
                    txtAttach.Text += string.Format("Error record not imported format error ") + Environment.NewLine;
                }
            }



            DateTime dtInsertStarted = DateTime.Now;
            TimeSpan duration = DateTime.Now - dtInsertStarted;
            txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", "", duration.Seconds, duration.Milliseconds) + Environment.NewLine;

            this.Cursor = Cursors.Arrow;
        }

        protected void CheckProcessPayFileRedo(String accFileName, String NewPath)
        {
            String[] fileSplit;
            int noFields;
            CultureInfo provider = CultureInfo.InvariantCulture;
            DateTime dateNew;

            int lineno = 0;

            fileSplit = accFileName.Split('_');
            String checkStr;
            String checkData;

            var textFieldParser = new TextFieldParser(NewPath + accFileName);

            textFieldParser.TextFieldType = FieldType.Delimited;
            textFieldParser.Delimiters = new[] { "|" };
            textFieldParser.TrimWhiteSpace = true;
            textFieldParser.HasFieldsEnclosedInQuotes = true;
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
            while (!textFieldParser.EndOfData)
            {

                lineno++;
                string[] parts = textFieldParser.ReadFields();
                noFields = parts.Length;

                txtProcess.Text = accFileName + " - " + noFields.ToString();

                var newPayFile = (from p in BdSuiteContext.ABSAPayFiles
                                  where p.Filename == accFileName.Replace(".zip", ".txt") && p.AccountNumber == parts[0] && p.Amount == Convert.ToDecimal(parts[1])
                                  && p.FinancialTransactionEffectiveDate == Convert.ToDateTime(parts[3]) && p.PaymentType == parts[5]
                                  select p).SingleOrDefault();

                newPayFile.Filename = accFileName.Replace(".zip", ".txt");
                newPayFile.AccType = fileSplit[3];
                dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                //   newPayFile.DateF = Convert.ToInt32(dateNew) + 2;
                newPayFile.FileDate = fileSplit[4];
                newPayFile.Book = fileSplit[2];

                checkData = parts[2];
                checkStr = checkData;

                if (noFields < 19)
                {

                    newPayFile.AccountNumber = parts[0];
                    newPayFile.Amount = Convert.ToDecimal(parts[1]);
                    if (!string.IsNullOrEmpty(parts[2]) && parts[2].Length > 9)
                    {
                        //  newPayFile.FinancialTransactionEffectiveDate = DateTime.ParseExact(parts[2], "yyyy/MM/dd", provider);
                    }
                    newPayFile.MemoCode = parts[3];
                    newPayFile.PaymentType = parts[4];
                    newPayFile.PaymentMethod = parts[5];
                }
                else
                {


                    newPayFile.Workgroup = parts[0];
                    newPayFile.AccountNumber = parts[1];
                    newPayFile.Amount = Convert.ToDecimal(parts[2]);
                    if (!string.IsNullOrEmpty(parts[3]))
                    {
                        newPayFile.FinancialTransactionEffectiveDate = Convert.ToDateTime(parts[3]);
                    }
                    newPayFile.MemoCode = parts[4];
                    newPayFile.PaymentType = parts[5];
                    newPayFile.PaymentMethod = parts[6];
                    // Tag6 = parts[7],
                    //  Tag6AssignmentDate = Convert.ToDateTime(parts[8]),
                    newPayFile.Tag4 = parts[9];
                    if (!string.IsNullOrEmpty(parts[10]))
                    {
                        newPayFile.Tag4AssignmentDate = Convert.ToDateTime(parts[10]);
                    }
                    newPayFile.Tag5 = parts[11];
                    if (!string.IsNullOrEmpty(parts[12]))
                    {
                        newPayFile.Tag5AssignmentDate = Convert.ToDateTime(parts[12]);
                    }
                    newPayFile.NAEDO = parts[13];
                    if (!string.IsNullOrEmpty(parts[14]))
                    {
                        newPayFile.Dateaccountopened = Convert.ToDateTime(parts[14]);
                    }
                    newPayFile.SiteName = parts[15];
                    newPayFile.IDNumber = parts[16];
                    newPayFile.Surname = parts[17];
                    newPayFile.SBUCode = parts[18];
                    newPayFile.Site = parts[19];
                }
                if (noFields >= 21)
                {
                    newPayFile.AttorneyCode = parts[20];
                    newPayFile.Commissionable = parts[21];
                }
                BdSuiteContext.SubmitChanges();

            }



            DateTime dtInsertStarted = DateTime.Now;

            TimeSpan duration = DateTime.Now - dtInsertStarted;
            txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", "", duration.Seconds, duration.Milliseconds) + Environment.NewLine;

            this.Cursor = Cursors.Arrow;
        }

        protected void CheckProcessRecallFileRedo(String accFileName, String NewPath)
        {
            String[] fileSplit;
            CultureInfo provider = CultureInfo.InvariantCulture;
            DateTime dateNew;
            fileSplit = accFileName.Split('_');
            int noFields;
            var textFieldParser = new TextFieldParser(NewPath + accFileName);

            textFieldParser.TextFieldType = FieldType.Delimited;
            textFieldParser.Delimiters = new[] { "|" };
            textFieldParser.TrimWhiteSpace = true;
            textFieldParser.HasFieldsEnclosedInQuotes = true;
            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();
            // Note: The [AccFile] table's Primary Key, ProductID, is an identity column, so we don't need to specify a value for each record we're about to create.
            while (!textFieldParser.EndOfData)
            {
                string[] parts = textFieldParser.ReadFields();
                noFields = parts.Length;
                var newRecallFile = (from p in BdSuiteContext.ABSARecallFiles
                                     where p.Filename == accFileName.Replace(".zip", ".txt") && p.Accountnumber == parts[1]
                                     select p).SingleOrDefault();

                newRecallFile.Filename = accFileName.Replace(".zip", ".txt");
                newRecallFile.AccType = fileSplit[3];
                newRecallFile.FileDate = fileSplit[4];
                dateNew = DateTime.ParseExact(fileSplit[4].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                //      newRecallFile.DateF = Convert.ToInt32(dateNew) + 2;
                if (noFields < 6)
                {
                    newRecallFile.Book = fileSplit[2];
                    newRecallFile.EDCWorkGroup = parts[0];
                    newRecallFile.Accountnumber = parts[1];
                    newRecallFile.LegalAccountNumber = parts[2];
                    newRecallFile.PrimaryPlasticNumber = parts[3];
                    newRecallFile.RecallDate = Convert.ToDateTime(parts[4]);
                }
                else
                {
                    newRecallFile.Book = fileSplit[2];
                    newRecallFile.EDCWorkGroup = parts[0];
                    newRecallFile.Accountnumber = parts[1];
                    newRecallFile.LegalAccountNumber = parts[2];
                    newRecallFile.PrimaryPlasticNumber = parts[3];
                    newRecallFile.RecallDate = Convert.ToDateTime(parts[4]);
                    newRecallFile.Recall_Area = parts[5];
                    newRecallFile.Prev_Risk = parts[6];
                    newRecallFile.Current_Risk = parts[7];


                }
                BdSuiteContext.SubmitChanges();

            }

            DateTime dtInsertStarted = DateTime.Now;

            //  Insert this List<> of records using a Bulk Insert


            TimeSpan duration = DateTime.Now - dtInsertStarted;
            txtAttach.Text += string.Format("Time to insert {0} records: {1}.{2} seconds", "", duration.Seconds, duration.Milliseconds) + Environment.NewLine;

            this.Cursor = Cursors.Arrow;
        }

        private void btnaddDataTable_Click(object sender, EventArgs e)
        {
            var table = new DataTable();

            var fileContents = File.ReadAllLines("yourFile");

            var splitFileContents = (from f in fileContents select f.Split(':')).ToArray();

            int maxLength = (from s in splitFileContents select s.Count()).Max();

            for (int i = 0; i < maxLength; i++)
            {
                table.Columns.Add();
            }

            foreach (var line in splitFileContents)
            {
                DataRow row = table.NewRow();
                row.ItemArray = (object[])line;
                table.Rows.Add(row);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            String itemAcc;

            itemAcc = nList[0];

            DataClassBDSuiteDataContext BdSuiteContext = new DataClassBDSuiteDataContext();

            var r = (from p in BdSuiteContext.AbsaFileLogs
                     where p.Filename == itemAcc
                     select p).SingleOrDefault();

            r.Processed = 0;
            BdSuiteContext.SubmitChanges();

            lstBxAccount.Items.Clear();
            nList.Clear();
            var q = from p in BdSuiteContext.AbsaFileLogs
                    where p.Processed == 0
                    orderby p.FileDate ascending, p.Book ascending
                    select p;
            foreach (var log in q)
            {

                lstBxAccount.Items.Add(log.Filename);
                nList.Add(log.Filename);

            }
            MessageBox.Show("Completed Moved Redo to Unprocessed");
        }

   
    }


}

    