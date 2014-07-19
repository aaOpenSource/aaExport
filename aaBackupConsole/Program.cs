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
        /// Calculate the correct object list based on multiple arguments
        /// </summary>
        /// <returns></returns>
        private static string FilterObjectList(string ObjectList)
        {
            string returnValue;
            string workingList;
            
            // Initalize the return val
            returnValue = "";

            // First set the object list to the passed object list
            workingList = ObjectList;

            // Now consider if the user passed any arguments that may later the object list
            
            // Key concept is that these functionss are filters, not adders.  So, if you run both filters it will reduce the list, never grow it

            // Change Log Timestamp Filter.. Basically Changes since the passed time
            if (_ChangeLogTimestampStartFilter !="")
            {
                returnValue = GetObjectListForChangeLogAllObjectsAfterTimestampAsCSV(DateTime.Parse(_ChangeLogTimestampStartFilter), ETemplateOrInstance.Both, workingList);
            }

            //Custom SQL
            if (_ChangeLogTimestampStartFilter != "")
            {
                returnValue = GetObjectListFromCustomSQL(_CustomSQLSelection, ETemplateOrInstance.Both, workingList);
            }

            log.Debug(returnValue);
            return returnValue;
        }

        private static IgObjects FilterGalaxyObjects(IgObjects GalaxyObjects)
        {
            //ArrayList galaxyObjectList = new ArrayList();

            List<String> galaxyObjectList = new List<String>();
            IgObjects returnGalaxyObjects;
            string workingList = "";
            string[] ObjectArray;
            
            // Get all of the items in the list.  Only good way to do this is 
            foreach (IgObject GObject in GalaxyObjects)
            {
                galaxyObjectList.Add(GObject.Tagname);
            }

            // Filter the Original List considering the Timestamp Filter
            if (_ChangeLogTimestampStartFilter != "")
            {
                workingList = GetObjectListForChangeLogAllObjectsAfterTimestampAsCSV(DateTime.Parse(_ChangeLogTimestampStartFilter), ETemplateOrInstance.Both, String.Join(",",galaxyObjectList.ToArray()));
            }

            //Trim the leading and trailing "
            workingList = workingList.Trim('"');

            // Split the working list in an array of string
            ObjectArray = workingList.Split(',');

            // Get an empty set of objects to start working with
            returnGalaxyObjects = GetEmptyIgObjects();

            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate,ObjectArray));
            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ObjectArray));

            // Check for any failures
            if ((returnGalaxyObjects == null) || (_Galaxy.CommandResult.Successful != true) || returnGalaxyObjects.count == 0)
            {
                // Failed to retrieve any objects from the query
                throw new Exception("Failed to retrieve objects to export.");
            }

            return returnGalaxyObjects;

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
                        return BackupObjectsToFile(EExportType.exportAsPDF, FilterObjectList(_ObjectList), _BackupFileName);

                    //Call the objects CSV Backup routine
                    case "ObjectsSingleCSV":
                        return BackupObjectsToFile(EExportType.exportAsCSV, FilterObjectList(_ObjectList), _BackupFileName);

                    //Exporting all Separate objects into AAPKG's
                    case "ObjectsSeparateAAPKG":
                        return BackupObjectsToFolder(EExportType.exportAsPDF, FilterObjectList(_ObjectList), _BackupFolderName);

                    //Exporting all Separate objects into CSV's
                    case "ObjectsSeparateCSV":
                        return BackupObjectsToFolder(EExportType.exportAsCSV, FilterObjectList(_ObjectList), _BackupFolderName);

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
                throw ex;
            }

        }

        #endregion

        #region SQL Data

        private static DataTable GetSQLData(string SQLQuery)
        {
            DataTable returnDataTable;

            try
            {

                // Check the connection
                if (_SQLConn.State != ConnectionState.Open)
                {
                    _SQLConn.ConnectionString = GetSQLConnectionString();
                    _SQLConn.Open();
                }

                if (_SQLConn.State != ConnectionState.Open)
                {
                    throw new Exception("SQL Connection Failed to Open");
                }

                // Setup our command
                SqlCommand _sqlCmd = new SqlCommand();

                _sqlCmd.Connection = _SQLConn;

                // Set the query text
                _sqlCmd.CommandType = CommandType.Text;
                _sqlCmd.CommandText = SQLQuery;

                // Execute
                returnDataTable = new DataTable("Data");
                returnDataTable.Load(_sqlCmd.ExecuteReader());

                return returnDataTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static string GetSQLConnectionString()
        {
            // Build the basic connection string
            // For now we force you to use a trusted connection just becuase it is good security hygiene.  Maybe later
            // we will consider letting the caller passs 

            return "Server=" + _GRNodeName + ";Database=" + _GalaxyName + ";Trusted_Connection=True;";
        }

        static string GetSQLForChangeLogAllObjectsAfterTimestampAsCSV(DateTime TargetTimestamp, ETemplateOrInstance ItemTypeSelection = ETemplateOrInstance.Both, string ObjectList = "")
        {

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT DISTINCT ',' + Tag_Name ");
            sb.Append(" ");
            sb.Append("FROM         dbo.gobject INNER JOIN");
            sb.Append(" ");
            sb.Append("dbo.gobject_change_log ON dbo.gobject.gobject_id = dbo.gobject_change_log.gobject_id INNER JOIN");
            sb.Append(" ");
            sb.Append("dbo.lookup_operation ON dbo.gobject_change_log.operation_id = dbo.lookup_operation.operation_id");
            sb.Append(" ");
            sb.Append(" Where");
            sb.Append(" ");
            sb.Append("Change_Date >='" + TargetTimestamp.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'");
            sb.Append(" ");
            sb.Append("and Operation_Name in ('CheckInSuccess','CreateInstance')");
            sb.Append(" ");

            // Add the clauses to limit by template or isntance
            switch (ItemTypeSelection)
            {
                case ETemplateOrInstance.Instance:
                    sb.Append("and is_Template=0");
                    sb.Append(" ");
                    break;
                case ETemplateOrInstance.Template:
                    sb.Append("and is_Template=1");
                    sb.Append(" ");
                    break;
                default:
                    //Do Nothing
                    break;
            }

            // Consider if the user passed an object list. If so, then use that to filter the results before return
            if (ObjectList != "")
            {
                // Fix up the passed CSV to make SQL like it a little better

                // Add ' at beginning and end
                ObjectList = "'" + ObjectList + "'";
                // Add ' around each ,
                ObjectList = ObjectList.Replace(",", "','");
                // Remove rogue spaces.  Tagnames will never have spaces
                ObjectList = ObjectList.Replace(" ", "");
                sb.Append("and Tag_Name in (" + ObjectList + ")");
                sb.Append(" ");
            }

            sb.Append("FOR XML PATH('')");

            log.Info(sb.ToString());

            return sb.ToString();

        }

        static string GetObjectListForChangeLogAllObjectsAfterTimestampAsCSV(DateTime TargetTimestamp, ETemplateOrInstance ItemTypeSelection = ETemplateOrInstance.Both, string ObjectList = "")
        {
            DataTable dt;
            string returnList;

            // First get the datatable by a SQL query
            dt = GetSQLData(GetSQLForChangeLogAllObjectsAfterTimestampAsCSV(TargetTimestamp, ItemTypeSelection, ObjectList));
            
            // If we have more than one row then fix up the string format
            if (dt.Rows.Count > 0)
            {
                returnList = dt.Rows[0][0].ToString();
                returnList = "\"" + returnList.Substring(1, returnList.Length - 1) + "\"";
            }
            else
            {
                returnList = "";
            }

            log.Debug(returnList);
            return returnList;
        }

        static string GetObjectListFromCustomSQL(string SQL, ETemplateOrInstance ItemTypeSelection = ETemplateOrInstance.Both, string ObjectList = "")
        {
            DataTable dt;
            string returnList;

            // First get the datatable by a SQL query
            dt = GetSQLData(SQL);

            // If we have more than one row then fix up the string format
            if (dt.Rows.Count > 0)
            {
                returnList = dt.Rows[0][0].ToString();
                returnList = "\"" + returnList.Substring(1, returnList.Length - 1) + "\"";
            }
            else
            {
                returnList = "";
            }

            log.Debug(returnList);
            return returnList;

        }

        static int exectest()

        {
            //DataTable dt;
            //string val;

            //dt = GetSQLData(GetSQLForChangeLogAllObjectsAfterTimestampAsCSV(new System.DateTime(2013, 01, 01), ETemplateOrInstance.Instance));
            //val = dt.Rows[0][0].ToString();
            //dt = GetSQLData(GetSQLForChangeLogAllObjectsAfterTimestampAsCSV(new System.DateTime(2013, 01, 01), ETemplateOrInstance.Template));
            //val = dt.Rows[0][0].ToString();
            //val = "\"" + val.Substring(1, val.Length - 1) + "\"";
            //log.Info(val);


            //dt = GetSQLData(GetSQLForChangeLogAllObjectsAfterTimestampAsCSV(new System.DateTime(2013, 01, 01), ETemplateOrInstance.Template,"AlarmOnHigh_001,AlarmOnLow_001,AlarmOnLow_002"));
            //if (dt.Rows.Count > 0)
            //{ 
            //val = dt.Rows[0][0].ToString();
            //val = "\"" + val.Substring(1, val.Length - 1) + "\"";
            //log.Info(val);
            //}
            //else
            //{ log.Info("No Rows"); }


            //val = "\"" + val.Substring(1, val.Length - 1) + "\"";
            //log.Info(val);

            //dt = GetSQLData(GetSQLForChangeLogAllObjectsAfterTimestampAsCSV(new System.DateTime(2013, 01, 01), ETemplateOrInstance.Both));
            //val = dt.Rows[0][0].ToString();
            //val = "\"" + val.Substring(1, val.Length - 1) + "\"";
            //log.Info(val);

            //dt = GetSQLData(GetSQLForChangeLogAllObjectsAfterTimestampAsCSV(new System.DateTime(2013, 01, 01)));
            //val = dt.Rows[0][0].ToString();
            //val = "\"" + val.Substring(1, val.Length - 1) + "\"";
            //log.Info(val);

            return 0;
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
            try
            {
                log.Debug("Checking and correcting filename " + BackupFileName + " to use .AAPKG");
                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".AAPKG");

                log.Info("Starting AAPKG Backup to " + BackupFileName);

                // Call complete backup routine
                _Galaxy.ExportAll(EExportType.exportAsPDF, BackupFileName);

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
            try
            {
                log.Debug("Checking and correcting filename " + BackupFileName + " to use .CSV");
                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".CSV");

                log.Info("Starting CSV Backup to " + BackupFileName);

                // Call complete backup routine
                _Galaxy.ExportAll(EExportType.exportAsCSV, BackupFileName);

                if (_Galaxy.CommandResult.Successful != true)
                {
                    // Failed to retrieve any objects from the query
                    log.Error("Error while executing BackupCompleteCSV");
                    return -2;
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
        /// Perform object backups generating a single file
        /// </summary>
        /// <returns></returns>
        private static int BackupObjectsToFile(EExportType ExportType, String ObjectList, String BackupFileName)
        {
            string[] ObjectArray;
            IgObjects GalaxyObjects;
            
            try
            {
                log.Info("Starting Objects Backup to " + _BackupFileName);

                // If the returned length is ok then stuff the objects into an array
                if (ObjectList.Length <= 0)
                {
                    // Object List not Long Enough
                    throw new Exception("Object list length = 0");
                }

                // Take the comma Separated values and split them into an array
                ObjectArray = ObjectList.Split(',');

                log.Debug("QueryObjectsByName for Templates");
                // Now get the template Objects into a GObjects set
                GalaxyObjects = _Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ref ObjectArray);

                if (_Galaxy.CommandResult.Successful != true)
                {
                    // Failed to retrieve any objects from the query
                    throw new Exception("Error while querying templates by tagname");
                }

                log.Debug("QueryObjectsByName for Instances");
                // Get Instance Objects.  We have to do this in two steps b/c we can't query templates and instances at the same time
                GalaxyObjects.AddFromCollection(_Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref ObjectArray));

                log.Debug("Verify GalaxyObject <> Null, Galaxy Command Success, and Galaxy Objet Count > 0");
                if ((GalaxyObjects == null) || (_Galaxy.CommandResult.Successful != true) || (GalaxyObjects.count == 0))
                {
                    // Failed to retrieve any objects from the query
                    throw new Exception("Failed to retrieve objects to export.");
                }

                log.Debug("Calling the BackupToFile function");
                // Perform Backup of the Objects in the Group
                if (BackupToFile(ExportType, GalaxyObjects, BackupFileName) != 0)
                {
                    throw new Exception("Error executing BackupToFile");
                }

                // Success
                return 0;

            }
            catch (Exception ex)
            {
                throw ex;              
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
                        log.Error("Failed to retrieve " + Item + " to export.");
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
                throw ex;                
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
                    throw new Exception("Failed to retrieve objects to export.");
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
                GalaxyObjects = FilterGalaxyObjects(GalaxyObjects);

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
                GalaxyObjects = FilterGalaxyObjects(GalaxyObjects);

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
                log.Error(ex.ToString());
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
