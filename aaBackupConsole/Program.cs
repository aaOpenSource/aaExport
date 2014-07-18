using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using ArchestrA.GRAccess;
//using logger;
using aaEncryption;
using log4net;

namespace aaBackupConsole
{
    class Program
    {
        // First things first, setup logging 
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// GR Access Class
        /// Used for accessing GR functionality
        /// </summary>
        static GRAccessApp _GRAccess;
        static IGalaxy _Galaxy;
        static IGalaxies _Galaxies;

        // Parameters passed by command line argument or file

        /// <summary>
        /// GR Hostname
        /// </summary>
        static string _GRNodeName;
        static string _GalaxyName;
        static string _Username;
        static string _Password;
        static string _BackupFileName;
        static string _BackupFolderName;
        static string _ObjectList;
        static string _BackupType;
        static CommandLine.Utility.Arguments _args;
        static string _IncludeConfigVersion;
        static string _FilterType;
        static string _Filter;
        static string _PasswordToEncrypt;
        static string _EncryptedPassword;
        
        static void Main(string[] args)
        {
        	
            try
            {
                // Start with the logging
                log4net.Config.BasicConfigurator.Configure();

                log.Info("Starting aaBackup");

                // First store off the arguments
                _args = new CommandLine.Utility.Arguments(args);
                
                // First call the setup routine
                if (Setup() != 1)
                {
                    log.Error("Setup Failed");
                    return;
                }

                //Console.WriteLine("Parsing Arguments");

                // Parse the input parameters
                if (ParseArguments(args) != 0)
                {
                    Console.WriteLine("Parsing Arguments Failed");
                    return;
                }

                // If the user has passed us a password to encrypt then do that and  bail out.
                if (_PasswordToEncrypt.Length > 0)
                {
                    Console.WriteLine("Encrypting Password");
                    WriteEncryptedPassword(_PasswordToEncrypt);
                    Console.WriteLine("Encryption Complete");
                    return;
                }

                // If the user has passed an encrypted password then decrypt it 
                // and set it to the current working password
                if (_EncryptedPassword.Length > 0)
                {
                    Console.WriteLine("Decrypted Password");
                    Console.WriteLine(DecryptPassword(_EncryptedPassword));

                    // Set the password to the the decrypted password
                    _Password = DecryptPassword(_EncryptedPassword);
                }

                //Console.WriteLine("Enter to Continue to Login");

                // Attempt to Connect
                if (Connect() != 0)
                {
                    Console.WriteLine("Connect Failed");
                    return;
                }

                if (PerformBackup(_BackupType) != 0)
                {
                    Console.WriteLine("Backup Failed");
                }

                Console.WriteLine("Enter to Continue");
                Console.ReadLine();

            }
            catch (Exception ex)
            {
                // Log the error
                //a2logger.LogError(ex.Message);              
                Console.Write(ex.Message.ToString());
                return;
            }
            finally
            {
                // Dispose of the Galaxy and GR Access Objects
                if(_Galaxy != null)
                {
                    _Galaxy = null;
                }

                if (_GRAccess != null)
                {
                    _GRAccess = null;
                }

                //Console.WriteLine("Enter to Continue to Finish");
                //Console.ReadLine();
            }
        }

        private static int Setup()
        {
            try
            {
                // Set the application identity
                //a2logger.LogSetIdentityName("aaBackupConsole");
                
                // Instantiate the GR Access App
                _GRAccess = new GRAccessApp();

                // Return success code
                return 0;
            }
            catch (Exception ex)
            {
                // Log the Error
                //a2logger.LogError(ex.Message);
                Console.Write(ex.Message.ToString());

                // Return an error code
                return -1;
            }
        }

        /// <summary>
        /// Process the arguments passed to the console program
        /// </summary>
        /// <param name="args"></param>
        private static int ParseArguments(string[] args)
        {
            try
            {
                //Parse the Command Line
                CommandLine.Utility.Arguments CommandLine = new CommandLine.Utility.Arguments(args);
                
                // Verify parameters passed are legal then stuff into variables
                if (CheckAndSetParameters(ref _GRNodeName, "GRNodeName", CommandLine,true,"localhost") != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _GalaxyName, "GalaxyName", CommandLine, true) != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _Username, "Username", CommandLine, true,"") != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _Password, "Password", CommandLine, true) != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _BackupFileName, "BackupFileName", CommandLine, true) != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _BackupType, "BackupType", CommandLine, true, "CompleteCAB") != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _BackupFolderName, "BackupFolder", CommandLine, true) != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _ObjectList, "ObjectList", CommandLine, true) != 0)
                {
                    return -2;
                }
                // Set to default if return is blank
                //if (_BackupType == "")
                //{
                //    _BackupType = "CompleteCAB";
                //}

/*                
				NOT USED - TODO: Need to figure out what the intended purpose of this was!
				
				if (CheckAndSetParameters(ref _FileDetail, "FileDetail", CommandLine, true) != 0)
                {
                    return -2;
                }
*/

                if (CheckAndSetParameters(ref _IncludeConfigVersion, "IncludeConfigVersion", CommandLine, true, "false") != 0)
                {
                    return -2;
                }

                //if (_IncludeConfigVersion == "")
                //{
                //    _IncludeConfigVersion = "false";
                //}

                if (CheckAndSetParameters(ref _FilterType, "FilterType", CommandLine, true) != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _Filter, "Filter", CommandLine, true) != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _PasswordToEncrypt, "PasswordToEncrypt", CommandLine, true) != 0)
                {
                    return -2;
                }

                if (CheckAndSetParameters(ref _EncryptedPassword, "EncryptedPassword", CommandLine, true) != 0)
                {
                    return -2;
                }


                return 0;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                return -1;
            }
        }

        /// <summary>
        /// Review the command line parameters for a specific parameter
        /// If it is present then push the parameter into the provided variable
        /// </summary>
        /// <param name="ParameterVariable"></param>
        /// <param name="ParameterName"></param>
        /// <param name="CommandLine"></param>
        /// <returns></returns>
        private static int CheckAndSetParameters(ref string ParameterVariable, string ParameterName, CommandLine.Utility.Arguments CommandLine, Boolean AllowEmpty, string DefaultValue = "") 
        {
            try
            {
                // Verify the parameter is present
                if (CommandLine[ParameterName] != null)
                {
                    // Set the variable if present
                    ParameterVariable = CommandLine[ParameterName].ToString();
                }
                else
                {
                    // If we are not allowing empties, then error.
                    if (!AllowEmpty)
                    {
                        // Warn the user the parameter is missing
                        Console.WriteLine("Missing parameter value for " + ParameterName);

                        // Return an error code
                        return -2;
                    }
                    else
                    {
                        // Set the return to an empty string
                        ParameterVariable = DefaultValue;
                    }
                }
                
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                return -1;
            }
        }

        /// <summary>
        /// Encrypt a password and write out to a file the display to the user
        /// </summary>
        /// <param name="PasswordToEncrypt"></param>
        private static void WriteEncryptedPassword(String PasswordToEncrypt)
        {
            cSimpleAES AES;
            String EncryptedPassword;
            String RandomFileName = "";

            try
            {

				// Create the AES object using the local machine fingerprint and static vector
				// This is necessary because we want to be able to reuse the encrypted password on the same machine				
            	AES = new cSimpleAES(cSecurity.FingerPrint.Value(),false);
            	
                // Encrypt the passed value
                EncryptedPassword = AES.EncryptToString(PasswordToEncrypt);

                // Get a random file name.  Just a little misdirection
                RandomFileName = System.IO.Path.GetTempFileName();

                // Write the password to the file
                System.IO.File.WriteAllText(RandomFileName, EncryptedPassword);

                //Run Notepad to show the results to the user so they can copy and paste it.
                System.Diagnostics.Process.Start("notepad.exe", RandomFileName);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                // Now Delete the file so it can't be picked up by something else.
                // Do this in the finally routine to guarantee it executes.
                System.IO.File.Delete(RandomFileName);
            }

        }

        /// <summary>
        /// Decrypt an encrypted password
        /// </summary>
        /// <param name="Encrypted"></param>
        /// <returns></returns>
        private static string DecryptPassword(String EncryptedPassword)
        {
            cSimpleAES AES;

            try
            {
				// Create the AES object using the local machine fingerprint and static vector
				// This is necessary because we want to be able to reuse the encrypted password on the same machine				
            	AES = new cSimpleAES(cSecurity.FingerPrint.Value(),false);

                // Return the decrypted string
                return AES.DecryptString(EncryptedPassword);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return "";
            }


        }


        /// <summary>
        /// Attempt to make a connection to the Galaxy
        /// </summary>
        /// <returns></returns>
        private static int Connect()
        {
            try
            {
                // Get a list of the available galaxies
                _Galaxies = _GRAccess.QueryGalaxies(_GRNodeName);

                //Get a reference to the Galaxy
                _Galaxy = _Galaxies[_GalaxyName];

                // Check to make sure we have a good reference to the Galaxy
                if (_Galaxy == null || !_GRAccess.CommandResult.Successful)
                {
                    Console.WriteLine(_GalaxyName + " is not a legal Galaxy.");
                    return -3;
                }

                // Attemp to Login
                _Galaxy.Login(_Username, _Password);
                //_Galaxy.Login("", "");
                //TODO:Revisit login b/c all logins seems to pass
                
                //Check for success
                if (!_Galaxy.CommandResult.Successful)
                {
                    Console.WriteLine("Galaxy Login Failed");
                    return -2;
                }

                // If we made it to hear we're good!

                Console.WriteLine("Login Succeeded");

                return 0;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                return -1;
            }
        }

        /// <summary>
        /// Perform the actual backup, switchin on the type of backup
        /// </summary>
        /// <returns></returns>
        private static int PerformBackup(string BackupType)
        {
            try
            {
                // Determine which type of backup and call the appropriate routine
                switch (BackupType)
                {
                    // Call the complete backup routine to generate a CAB
                    case "CompleteCAB":
                        return BackupCompleteCAB(_BackupFileName);

                        // Call the complete AAPKG backup routine
                    case "CompleteAAPKG":
                        return BackupCompleteAAPKG(_BackupFileName);

                    // Call the complete CSV backup routine
                    case "CompleteCSV":
                        return BackupCompleteCSV(_BackupFileName);

                        //Call the objects AAPKG Backup routine
                    case "ObjectsSingleAAPKG":
                        return BackupObjectsToFile(EExportType.exportAsPDF, _ObjectList, _BackupFileName);

                    //Call the objects CSV Backup routine
                    case "ObjectsSingleCSV":
                        return BackupObjectsToFile(EExportType.exportAsCSV, _ObjectList, _BackupFileName);

                    //Exporting all Separate objects into AAPKG's
                    case "ObjectsSeparateAAPKG":
                        return BackupObjectsToFolder(EExportType.exportAsPDF, _ObjectList,_BackupFolderName);

                    //Exporting all Separate objects into CSV's
                    case "ObjectsSeparateCSV":
                        return BackupObjectsToFolder(EExportType.exportAsCSV,_ObjectList,_BackupFolderName);

                    //Export All Templates to an Single AAPKG
                    case "AllTemplatesAAPKG":
                        return BackupBySingleFilter(EExportType.exportAsPDF, _BackupFileName,"", "namedLike", "%", ETemplateOrInstance.Template);

                    //Export all Instances to Single AAPKG
                    case "AllInstancesAAPKG":
                        return BackupBySingleFilter(EExportType.exportAsPDF, _BackupFileName, "", "namedLike", "%", ETemplateOrInstance.Instance);

                    //Export all Instances to Single CSV
                    case "AllInstancesCSV":
                        return BackupBySingleFilter(EExportType.exportAsCSV, _BackupFileName, "", "namedLike", "%", ETemplateOrInstance.Instance);

                    //Export All Templates to Separate AAPKG Files
                    case "AllTemplatesSeparateAAPKG":
                        return BackupBySingleFilter(EExportType.exportAsPDF, "", _BackupFolderName, "namedLike", "%", ETemplateOrInstance.Template);

                    //Export all Instances to Separate AAPKG
                    case "AllInstancesSeparateAAPKG":
                        return BackupBySingleFilter(EExportType.exportAsPDF, "", _BackupFolderName, "namedLike", "%", ETemplateOrInstance.Instance);

                    //Export all Instances to Separate AAPKG
                    case "AllInstancesSeparateCSV":
                        return BackupBySingleFilter(EExportType.exportAsCSV, "", _BackupFolderName, "namedLike", "%", ETemplateOrInstance.Instance);

                    //Export Objects Based on Filter Criteria to single AAPKG
                    case "FilteredObjectsAAPKG":
                        return BackupBySingleFilter(EExportType.exportAsPDF, _BackupFileName, "", _FilterType , _Filter, ETemplateOrInstance.Both);

                    //Export Objects Based on Filter Criteria to single CSV
                    case "FilteredObjectsCSV":
                        return BackupBySingleFilter(EExportType.exportAsCSV, _BackupFileName, "", _FilterType, _Filter, ETemplateOrInstance.Both);

                    //Export Objects Based on Filter Criteria to Separate AAPKG
                    case "FilteredObjectsSeparateAAPKG":
                        return BackupBySingleFilter(EExportType.exportAsPDF, "", _BackupFolderName, _FilterType, _Filter, ETemplateOrInstance.Both);

                    //Export Objects Based on Filter Criteria to Separate CSV
                    case "FilteredObjectsSeparateCSV":
                        return BackupBySingleFilter(EExportType.exportAsCSV, "", _BackupFolderName, _FilterType, _Filter, ETemplateOrInstance.Both);

                    default:
                        break;
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                return -1;
            }

        }

#region "Backup Routines"

        /// <summary>
        /// Perform a complete backup generating a CAB
        /// </summary>
        /// <returns></returns>
        private static int BackupCompleteCAB(String BackupFileName)
        {
            int ProcessId;

            try
            {
                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".CAB");

                // Get the current PID
                ProcessId = System.Diagnostics.Process.GetCurrentProcess().Id;

                if (ProcessId == 0)
                {
                    Console.WriteLine("Inavlid ProcessID");
                    return -2;
                }

                Console.WriteLine("Starting CAB Backup to " + BackupFileName);

                // Call complete backup routine
                _Galaxy.Backup(ProcessId, BackupFileName, _GRNodeName, _GalaxyName);
                
                Console.WriteLine("Backup CAB Complete");

                return 0;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                return -1;
            }

        }

        /// <summary>
        /// Perform a complete backup generating an AAPKG
        /// </summary>
        /// <returns></returns>
        private static int BackupCompleteAAPKG(String BackupFileName)
        {
            try
            {

                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".AAPKG");

                Console.WriteLine("Starting AAPKG Backup to " + BackupFileName);

                // Call complete backup routine
                _Galaxy.ExportAll(EExportType.exportAsPDF, BackupFileName);

                Console.WriteLine("Backup AAPKG Complete");

                return 0;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                return -1;
            }
        }

        /// <summary>
        /// Perform a complete backup generating a CSV
        /// </summary>
        /// <returns></returns>
        private static int BackupCompleteCSV(String BackupFileName)
        {
            try
            {
                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".CSV");

                Console.WriteLine("Starting CSV Backup to " + BackupFileName);

                // Call complete backup routine
                _Galaxy.ExportAll(EExportType.exportAsCSV, BackupFileName);

                Console.WriteLine("Backup CSV Complete");

                return 0;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                return -1;
            }
        }

        /// <summary>
        /// Perform object backups generating a single file
        /// </summary>
        /// <returns></returns>
        private static int BackupObjectsToFile(EExportType ExportType, String ObjectList, String BackupFileName)
        {
            string[] ObjectArray;
            IgObjects GalaxyObjects;
            
            try
            {
                Console.WriteLine("Starting Objects Backup to " + _BackupFileName);

                // If the returned length is ok then stuff the objects into an array
                if (ObjectList.Length == 0)
                {
                    // Object List not Long Enough
                    return -3;
                }

                // Take the comma Separated values and split them into an array
                ObjectArray = ObjectList.Split(',');

                // Now get the template Objects into a GObjects set
                GalaxyObjects = _Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ref ObjectArray);

                // Get Instance Objects.  We have to do this in two steps b/c we can't query templates and instances at the same time
                GalaxyObjects.AddFromCollection(_Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref ObjectArray));

                if ((GalaxyObjects == null) || (_Galaxy.CommandResult.Successful != true) || (GalaxyObjects.count == 0))
                {
                    // Failed to retrieve any objects from the query
                    Console.WriteLine("Failed to retrieve objects to export.");
                    return -4;
                }

                

                // Perform Backup of the Objects in the Group
                BackupToFile(ExportType, GalaxyObjects, BackupFileName);

                return 0;

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                return -1;
            }
        }

        /// <summary>
        /// Perform object backups generating Separate files
        /// </summary>
        /// <returns></returns>
        private static int BackupObjectsToFolder(EExportType ExportType, String ObjectList, String BackupFolderName)
        {
            string[] ObjectArray;
            IgObjects GalaxyObjects;

            string[] SingleObjectItem;

            try
            {   
                // Instantiate the single Object item array
                SingleObjectItem = new string[1];



                // If the returned length is ok then stuff the objects into an array
                if (ObjectList.Length == 0)
                {
                    // Object List not Long Enough
                    return -5;
                }

                // Take the comma Separated values and split them into an array
                ObjectArray = ObjectList.Split(',');

                // Loop through all the objects and export them
                foreach (String Item in ObjectArray)
                {
                    // Copy itemname in array
                    SingleObjectItem[0] = Item;

                    // Figure out if it's a Template or Instance
                    if (SingleObjectItem[0].Substring(0, 1) == "$")
                    {
                        //Template
                        // Get the Objects Reference via Query
                        GalaxyObjects = _Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ref SingleObjectItem);
                    }
                    else
                    {
                        //Instance
                        // Get the Objects Reference via Query
                        GalaxyObjects = _Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref SingleObjectItem);
                    }

                    // Check for valid object list
                    if ((GalaxyObjects == null) || (_Galaxy.CommandResult.Successful != true) || (GalaxyObjects.count == 0))
                    {
                        // Failed to retrieve any objects from the query
                        Console.WriteLine("Failed to retrieve " + Item + " to export.");
                    }
                    else
                    {
                        // Object list is good, let's continue
                        return BackupToFolder(ExportType, GalaxyObjects, BackupFolderName);
                    }
                }

                return 0;

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                return -1;
            }
        }

        /// <summary>
        /// Perform object backups generating a single file from a single filter set
        /// </summary>
        /// <returns></returns>
        private static int BackupBySingleFilter(EExportType ExportType, String BackupFileName, String BackupFolder, String FilterType, String Filter, ETemplateOrInstance TemplateOrInstance)
        {

            IgObjects GalaxyObjects;
            Boolean ExecBackupToSingleFile;
            Boolean ExecBackupToFolder;

            try
            {
                // Figure out if we're backing up to single file, folder or both
                ExecBackupToSingleFile = (BackupFileName.Length > 0);
                ExecBackupToFolder = (BackupFolder.Length > 0);

                // Get an empty set of objects to start working with
                GalaxyObjects = GetEmptyIgObjects();

                // Do we need to include templates?
                if((TemplateOrInstance == ETemplateOrInstance.Template) || (TemplateOrInstance == ETemplateOrInstance.Both))
                {
                    // Now get the template Objects into a GObjects set
                    GalaxyObjects.AddFromCollection(_Galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ConditionType(FilterType), (object)Filter,EMatch.MatchCondition));
                }

                // Do we need to include instances?
                if ((TemplateOrInstance == ETemplateOrInstance.Instance) || (TemplateOrInstance == ETemplateOrInstance.Both))
                {
                    // Now get the template Objects into a GObjects set
                    GalaxyObjects.AddFromCollection(_Galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsInstance, ConditionType(FilterType), (object)Filter, EMatch.MatchCondition));
                }

                if ((GalaxyObjects == null) || (_Galaxy.CommandResult.Successful != true) || GalaxyObjects.count == 0)
                {
                    // Failed to retrieve any objects from the query
                    Console.WriteLine("Failed to retrieve objects to export.");
                    return -4;
                }

                // If we are backing up all items to a single file then make a simple call to export all the objects
                if (ExecBackupToSingleFile)
                {
                    return BackupToFile(ExportType,GalaxyObjects,BackupFileName);
                }

                // If we are backing up to a folder that means we want multiple files
                if (ExecBackupToFolder)
                {
                    return BackupToFolder(ExportType, GalaxyObjects, BackupFolder);
                }

                return 0;

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                return -1;
            }
        }
        
        /// <summary>
        /// Backup GObjects to a Single File
        /// </summary>
        /// <param name="ExportType"></param>
        /// <param name="GalaxyObjects"></param>
        /// <param name="BackupFileName"></param>
        /// <returns></returns>
        private static int BackupToFile(EExportType ExportType, IgObjects GalaxyObjects ,String BackupFileName)
        {
            try
            {
                // Make sure we have the right extension on the backup file
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, CorrectExtension(ExportType));

                Console.WriteLine("Backing up to " + BackupFileName);

                // Perform the actual export
                GalaxyObjects.ExportObjects(ExportType, BackupFileName);

                if (GalaxyObjects.CommandResults.CompletelySuccessful != true)
                {
                    Console.WriteLine("Export not completely successful");
                    Console.WriteLine(GalaxyObjects.CommandResults.ToString());
                    return -2;
                }
                else
                {
                    Console.WriteLine("Backup to " + BackupFileName + " succeeded.");
                }

                return 0;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return -1;
            }
        }

        /// <summary>
        /// Backup a Single GObject to a File
        /// </summary>
        /// <param name="ExportType"></param>
        /// <param name="GalaxyObject"></param>
        /// <param name="BackupFileName"></param>
        /// <returns></returns>
        private static int BackupToFile(EExportType ExportType, IgObject GalaxyObject, String BackupFileName)
        {
            IgObjects GalaxyObjects;
            String[] SingleItemName  = new String[1];

            try
            {
                SingleItemName[0] = GalaxyObject.Tagname;

                if(SingleItemName[0].Substring(0,1) == "$")
                {
                    //Template
                    GalaxyObjects = _Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ref SingleItemName);
                }
                else
                {
                    // Instance
                    GalaxyObjects = _Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref SingleItemName);
                }

                if (GalaxyObjects.count > 0)
                {
                    return BackupToFile(ExportType, GalaxyObjects, BackupFileName);
                }
                else
                {
                    return -2;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return -1;
            }
        }

        /// <summary>
        /// Backup GObjects to multiple files in a single folder
        /// </summary>
        /// <param name="ExportType"></param>
        /// <param name="GalaxyObjects"></param>
        /// <param name="BackupFolderName"></param>
        /// <returns></returns>
        private static int BackupToFolder(EExportType ExportType, IgObjects GalaxyObjects, String BackupFolderName)
        {
            String[] SingleItemName = new String[1];
            IgObjects SingleObject;
            String BackupFileName;
            String Extension;
            String ConfigVersion;

            try
            {
                // Get Extension
                Extension = CorrectExtension(ExportType);

                //Create the Backup FOlder if it doesn't exist
                if (!System.IO.Directory.Exists(BackupFolderName))
                {
                    System.IO.Directory.CreateDirectory(BackupFolderName);
                }

                // Double check that we have the folder.  If it's missing then error out
                if (!System.IO.Directory.Exists(BackupFolderName))
                {
                    Console.WriteLine("Missing Directory " + BackupFolderName);
                    return -2;
                }

                //Need to iterate through the the objects
                foreach(IgObject Item in GalaxyObjects)
                {
                    // Populate the single item array with the item's tagname
                    SingleItemName[0] = Item.Tagname;

                    //Default Config Version Text to Blank
                    ConfigVersion = "";

                    //If we require config version in the filename then figure it out then add it
                    if (_IncludeConfigVersion == "true")
                    {
                        ConfigVersion = "-v" + Item.ConfigVersion.ToString();
                    }

                    // Calculate the appropriate backup filename
                    BackupFileName = BackupFolderName + "\\" + Item.Tagname + ConfigVersion + Extension;

                    // Figure out if we're dealing with a Template or Instance
                    if(Item.Tagname.Substring(0,1) == "$")
                    {
                        //Template
                        SingleObject = _Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ref SingleItemName);
                    }
                    else
                    {
                        // Instance
                        SingleObject = _Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref SingleItemName);
                    }

                    // Make sure we have an actual object in the collection
                    if (SingleObject.count > 0)
                    {
                        // Perform the actual backup
                        BackupToFile(ExportType, SingleObject, BackupFileName);
                    }
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return -1;
            }
        }

#endregion

        #region Utilities

        /// <summary>
        /// Return an empty set of GObjects so we can simply add to it
        /// </summary>
        /// <returns></returns>
        private static IgObjects GetEmptyIgObjects()
        {
            string[] DummyStringRef;
            try
            {
                DummyStringRef = new string[1];
                DummyStringRef[0] = "-";
                return _Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref DummyStringRef);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return null;
            }
        }

        /// <summary>
        /// Return an EConditionType when given a string that maps to the condition type 
        /// </summary>
        /// <param name="ConditionType"></param>
        /// <returns></returns>
        private static EConditionType ConditionType(String ConditionType)
        {
            // Just do a bigh switch case on all the different kinds, return 
            // the correct reference
            switch (ConditionType)
            {
            case ("derivedOrInstantiatedFrom"): return EConditionType.derivedOrInstantiatedFrom;
            case ("basedOn"): return EConditionType.basedOn;
            case ("containedBy"): return EConditionType.containedBy;
            case ("hostEngineIs"): return EConditionType.hostEngineIs;
            case ("belongsToArea"): return EConditionType.belongsToArea;
            case ("assignedTo"): return EConditionType.assignedTo;
            case ("withinSecurityGroup"): return EConditionType.withinSecurityGroup;
            case ("createdBy"): return EConditionType.createdBy;
            case ("lastModifiedBy"): return EConditionType.lastModifiedBy;
            case ("checkedOutBy"): return EConditionType.checkedOutBy;
            case ("namedLike"): return EConditionType.namedLike;
            case ("validationStatusIs"): return EConditionType.validationStatusIs;
            case ("deploymentStatusIs"): return EConditionType.deploymentStatusIs;
            case ("checkoutStatusIs"): return EConditionType.checkoutStatusIs;
            case ("objectCategoryIs"): return EConditionType.objectCategoryIs;
            case ("hierarchicalNameLike"): return EConditionType.hierarchicalNameLike;
            case ("NameEquals"): return EConditionType.NameEquals;
            case ("NameSpaceldls"): return EConditionType.NameSpaceIdIs;
            default: return 0;
            }
        }

        /// <summary>
        /// Calculate extension based on Export Type
        /// </summary>
        /// <param name="ExportType"></param>
        /// <returns></returns>
        private static string CorrectExtension(EExportType ExportType)
        {
            // Set the correct extension based on type of export
            if (ExportType == EExportType.exportAsPDF)
            {
                return ".AAPKG";
            }

            if (ExportType == EExportType.exportAsCSV)
            {
                return ".CSV";
            }

            return "";
        }

        #endregion

        #region Enums

        enum ETemplateOrInstance
        {
            Template = 1,
            Instance = 2,
            Both = 3
        }

		
        #endregion

    }
}
