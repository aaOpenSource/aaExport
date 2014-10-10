using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using ArchestrA.GRAccess;
using aaEncryption;
using log4net;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Threading;
using aaObjectSelection;
using System.Diagnostics;


namespace aaBackupConsole
{
    class Program
    {

        #region Declarations

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
        static string _IncludeConfigVersion;
        static string _FilterType;
        static string _Filter;
        static string _PasswordToEncrypt;
        static string _EncryptedPassword;
        static string _ChangeLogTimestampStartFilter;
        static string _CustomSQLSelection;
        
        static CommandLine.Utility.Arguments _args;
        static SqlConnection _SQLConn = new SqlConnection();

        #endregion

        #region Core

        static void Main(string[] args)
        {
        	
            try
            {
                // Start with the logging
                log4net.Config.BasicConfigurator.Configure();

                log.Info("Starting aaBackup");

                // First store off the arguments
                _args = new CommandLine.Utility.Arguments(args);

                // Parse the input parameters
                ParseArguments(args);

                // First call the setup routine
                Setup();

                // If the user has passed us a password to encrypt then do that and  bail out.
                if (_PasswordToEncrypt.Length > 0)
                {
                    log.Info("Encrypting password");
                    WriteEncryptedPassword(_PasswordToEncrypt);
                    log.Info("Password encryption complete");
                    return;
                }

                // If the user has passed an encrypted password then decrypt it 
                // and set it to the current working password
                if (_EncryptedPassword.Length > 0)
                {
                    log.Info("Decrypting password");
                    
                    // Set the password to the the decrypted password
                    _Password = DecryptPassword(_EncryptedPassword);
                }

                // Attempt to Connect
                Connect();

                // Execute the Backup
                PerformBackup(_BackupType);

            }
            catch (Exception ex)
            {
                log.Error(ex);
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
                
                Console.WriteLine("Enter to Continue to Finish");
                Console.ReadLine();
            }
        }

        private static int Setup()
        {
            try
            {                
                // Instantiate the GR Access App
                _GRAccess = new GRAccessApp();

                // Return success code
                return 0;
            }
            catch (Exception ex)
            {
                throw ex;
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
                CheckAndSetParameters(ref _GRNodeName, "GRNodeName", CommandLine,true,"localhost");
                CheckAndSetParameters(ref _GalaxyName, "GalaxyName", CommandLine, true);                
                CheckAndSetParameters(ref _Username, "Username", CommandLine, true,"");                
                CheckAndSetParameters(ref _Password, "Password", CommandLine, true);                
                CheckAndSetParameters(ref _BackupFileName, "BackupFileName", CommandLine, true);                
                CheckAndSetParameters(ref _BackupType, "BackupType", CommandLine, true, "CompleteCAB");                
                CheckAndSetParameters(ref _BackupFolderName, "BackupFolder", CommandLine, true);                
                CheckAndSetParameters(ref _ObjectList, "ObjectList", CommandLine, true);
                  
                /*
				NOT USED - TODO: Need to figure out what the intended purpose of this ws!				
				CheckAndSetParameters(ref _FileDetail, "FileDetail", CommandLine, true);
                */

                CheckAndSetParameters(ref _IncludeConfigVersion, "IncludeConfigVersion", CommandLine, true, "false");
                CheckAndSetParameters(ref _FilterType, "FilterType", CommandLine, true);
                CheckAndSetParameters(ref _Filter, "Filter", CommandLine, true);
                CheckAndSetParameters(ref _PasswordToEncrypt, "PasswordToEncrypt", CommandLine, true);
                CheckAndSetParameters(ref _EncryptedPassword, "EncryptedPassword", CommandLine, true);
                CheckAndSetParameters(ref _ChangeLogTimestampStartFilter, "ChangeLogTimestampStartFilter", CommandLine, true);
                CheckAndSetParameters(ref _CustomSQLSelection, "CustomSQLSelection", CommandLine, true);

                // Success
                return 0;
            }
            catch (Exception ex)
            {
                throw ex;
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
                log.Debug("Verifying parameter " + ParameterName + " is not null");
                // Verify the parameter is present
                if (CommandLine[ParameterName] != null)
                {
                    // Set the variable if present
                    ParameterVariable = CommandLine[ParameterName].ToString();
                    log.Debug("Set " + ParameterName + " to " + ParameterVariable);
                }
                else
                {
                    log.Debug("Considering allowempty");
                    // If we are not allowing empties, then error.
                    if (!AllowEmpty)
                    {
                        // Warn the user the parameter is missing
                        throw new Exception("Missing parameter value for " + ParameterName);
                    }
                    else
                    {
                        // Set the return to an empty string
                        ParameterVariable = DefaultValue;
                        log.Debug("Set " + ParameterName + " to default value of " + ParameterVariable);
                    }
                }

                log.Debug("Returning Success");
                // Success
                return 0;
            }
            catch (Exception ex)
            {
                throw ex;
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
                throw ex;
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
                throw ex;                
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

                log.Debug("Retrieving Galaxies for " + _GRNodeName);
                // Get a list of the available galaxies
                _Galaxies = _GRAccess.QueryGalaxies(_GRNodeName);

                log.Debug("Getting Galaxy Reference for " + _GalaxyName);
                //Get a reference to the Galaxy
                _Galaxy = _Galaxies[_GalaxyName];

                log.Debug("Checking for Success");
                // Check to make sure we have a good reference to the Galaxy
                if (_Galaxy == null || !_GRAccess.CommandResult.Successful)
                {
                    log.Error(_GalaxyName + " is not a legal Galaxy.");
                    return -3;
                }

                log.Debug("Logging in with Username " + _Username);
                // Attempt to Login
                _Galaxy.Login(_Username, _Password);

                log.Debug("Checking for login success");
                //Check for success
                if (!_Galaxy.CommandResult.Successful)
                {
                    log.Error("Galaxy Login Failed");
                    return -2;
                }

                // If we made it to hear we're good!
                log.Info("Login Succeeded");

                // Return control
                return 0;

            }
            catch (Exception ex)
            {
                throw ex;
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
                // Instantiate a new object selection class
                cObjectList objectList = new cObjectList(_Galaxy, _GRNodeName);

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
                        return BackupToFile(EExportType.exportAsPDF, objectList.GetObjectsFromStringList(_ObjectList, true), _BackupFileName);

                    //Call the objects CSV Backup routine
                    case "ObjectsSingleCSV":
                        return BackupToFile(EExportType.exportAsCSV, objectList.GetObjectsFromStringList(_ObjectList, true), _BackupFileName);

                    //Exporting all Separate objects into AAPKG's
                    case "ObjectsSeparateAAPKG":
                        return BackupToFolder(EExportType.exportAsPDF, objectList.GetObjectsFromStringList(_ObjectList, true), _BackupFolderName);

                    //Exporting all Separate objects into CSV's
                    case "ObjectsSeparateCSV":
                        return BackupToFolder(EExportType.exportAsCSV, objectList.GetObjectsFromStringList(_ObjectList, true), _BackupFolderName);

                    //Export All Templates to an Single AAPKG
                    case "AllTemplatesAAPKG":
                        return BackupToFile(EExportType.exportAsPDF, objectList.GetAllTemplates() ,_BackupFileName);

                    //Export all Instances to Single AAPKG
                    case "AllInstancesAAPKG":
                        return BackupToFile(EExportType.exportAsPDF, objectList.GetAllInstances(), _BackupFileName);

                    //Export all Instances to Single CSV
                    case "AllInstancesCSV":
                        return BackupToFile(EExportType.exportAsCSV, objectList.GetAllInstances(), _BackupFileName);

                    //Export All Templates to Separate AAPKG Files
                    case "AllTemplatesSeparateAAPKG":
                        return BackupToFolder(EExportType.exportAsPDF, objectList.GetAllTemplates(), _BackupFolderName);

                    //Export all Instances to Separate AAPKG
                    case "AllInstancesSeparateAAPKG":
                        return BackupToFolder(EExportType.exportAsPDF, objectList.GetAllInstances(), _BackupFolderName);                        

                    //Export all Instances to Separate CSV
                    case "AllInstancesSeparateCSV":
                        return BackupToFolder(EExportType.exportAsCSV, objectList.GetAllInstances(), _BackupFolderName);                        


                    //Export Objects Based on Filter Criteria to single AAPKG
                    case "FilteredObjectsAAPKG":
                        return BackupToFile(EExportType.exportAsPDF, objectList.GetObjectsFromSingleFilter(_FilterType,_Filter,aaObjectSelection.cObjectList.ETemplateOrInstance.Both,true), _BackupFileName);

                    //Export Objects Based on Filter Criteria to single CSV
                    case "FilteredObjectsCSV":
                        return BackupToFile(EExportType.exportAsCSV, objectList.GetObjectsFromSingleFilter(_FilterType, _Filter, aaObjectSelection.cObjectList.ETemplateOrInstance.Both, true), _BackupFileName);

                    //Export Objects Based on Filter Criteria to Separate AAPKG
                    case "FilteredObjectsSeparateAAPKG":
                        return BackupToFolder(EExportType.exportAsPDF, objectList.GetObjectsFromSingleFilter(_FilterType, _Filter, aaObjectSelection.cObjectList.ETemplateOrInstance.Both, true), _BackupFolderName);

                    //Export Objects Based on Filter Criteria to Separate CSV
                    case "FilteredObjectsSeparateCSV":
                        return BackupToFolder(EExportType.exportAsCSV, objectList.GetObjectsFromSingleFilter(_FilterType, _Filter, aaObjectSelection.cObjectList.ETemplateOrInstance.Both, true), _BackupFolderName);

                    default:
                        break;
                }

                return 0;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        #endregion

        #region Backup Routines

        /// <summary>
        /// Perform a complete backup generating a CAB
        /// </summary>
        /// <returns></returns>
        private static int BackupCompleteCAB(String BackupFileName)
        {
            int ProcessId;

            try
            {
                log.Debug("Checking and correcting filename " + BackupFileName + " to use .CAB");
                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".CAB");

                // Get the current PID
                ProcessId = System.Diagnostics.Process.GetCurrentProcess().Id;
                log.Debug("Got Process ID " + ProcessId.ToString());

                if (ProcessId == 0)
                {
                    log.Error("Inavlid ProcessID");
                    return -2;
                }

                log.Info("Starting CAB Backup to " + BackupFileName);

                // Call complete backup routine
                _Galaxy.Backup(ProcessId, BackupFileName, _GRNodeName, _GalaxyName);

                log.Info("Backup CAB Complete");
                
                // Success
                return 0;
            }
            catch (Exception ex)
            {
                throw ex;                
            }

        }

        /// <summary>
        /// Perform a complete backup generating an AAPKG
        /// </summary>
        /// <returns></returns>
        private static int BackupCompleteAAPKG(String BackupFileName)
        {
            //IgObjects workingGObjectList;
            cObjectList ObjectList = new cObjectList(_Galaxy, _GRNodeName);

            try
            {
                log.Debug("Checking and correcting filename " + BackupFileName + " to use .AAPKG");
                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".AAPKG");

                log.Info("Starting AAPKG Backup to " + BackupFileName);
                
                // Set the filter criteria
                ObjectList.ChangeLogTimestampStartFilter = _ChangeLogTimestampStartFilter;
                ObjectList.CustomSQLSelection = _CustomSQLSelection;

                //Perform the export
                ObjectList.GetCompleteObjectList(true).ExportObjects(EExportType.exportAsPDF, BackupFileName);
                //workingGObjectList.ExportObjects(EExportType.exportAsPDF, BackupFileName);

                if (_Galaxy.CommandResult.Successful != true)
                {
                    // Failed to retrieve any objects from the query
                    log.Error("Error while executing BackupCompleteAAPKG");
                }

                log.Info("Backup AAPKG Complete");

                return 0;
            }
            catch (Exception ex)
            {
                throw ex;                
            }
        }

        /// <summary>
        /// Perform a complete backup generating a CSV
        /// </summary>
        /// <returns></returns>
        private static int BackupCompleteCSV(String BackupFileName)
        {
            //IgObjects workingGObjectList;
            cObjectList ObjectList = new cObjectList(_Galaxy, _GRNodeName);

            try
            {
                log.Debug("Checking and correcting filename " + BackupFileName + " to use .CSV");
                
                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".CSV");

                log.Info("Starting CSV Backup to " + BackupFileName);

                // Set the filter criteria
                ObjectList.ChangeLogTimestampStartFilter = _ChangeLogTimestampStartFilter;
                ObjectList.CustomSQLSelection = _CustomSQLSelection;

                ObjectList.GetCompleteObjectList(true).ExportObjects(EExportType.exportAsCSV, BackupFileName);

                if (_Galaxy.CommandResult.Successful != true)
                {
                    // Failed to retrieve any objects from the query
                    log.Error("Error while executing BackupCompleteCSV");
                }

                log.Info("Backup CSV Complete");

                return 0;
            }
            catch (Exception ex)
            {
                throw ex;                
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

                log.Info("Backing up to " + BackupFileName);

                // Filter the Objects list
                //GalaxyObjects = FilterGalaxyObjects(GalaxyObjects);

                // Perform the actual export
                GalaxyObjects.ExportObjects(ExportType, BackupFileName);

                if (GalaxyObjects.CommandResults.CompletelySuccessful != true)
                {
                    throw new Exception(GalaxyObjects.CommandResults.ToString());
                }
                else
                {
                    log.Info("Backup to " + BackupFileName + " succeeded.");
                }

                return 0;
            }
            catch(Exception ex)
            {
                log.Error(ex.ToString());
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

                return 0;
            }
            catch (Exception ex)
            {
                throw ex;
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
                    throw new Exception("Missing Directory " + BackupFolderName);
                }

                // Filter the Objects list
                //GalaxyObjects = FilterGalaxyObjects(GalaxyObjects);

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
                log.Error(ex.ToString());
                return -1;
            }
        }

#endregion


        #region Utilities

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
    }
}
