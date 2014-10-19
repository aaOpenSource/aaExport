using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ArchestrA.GRAccess;

using Classes.ObjectList;
using Classes.GalaxyHelper;

namespace Classes.Export
{
    class caaExport
    {

        #region Declarations

        private IGalaxy _galaxy;
        private string _galaxyName = "";        
        private string _grNodeName = "";
        private string _username = "";
        private string _password = "";
        private string _backupFileName = "";
        private string _backupFolderName = "";
        private string _delimitedObjectList = "";
        private string _backupType = "";
        private bool _includeConfigVersion = false;
        private string _filterType = "";
        private string _filter = "";
        private DateTime _changeLogTimestampStartFilter = DateTime.Parse("1/1/1970");
        private string _customSQLSelection = "";
        private bool _overwriteFiles = true;
        private string _objectListFile = "";
        public EBackupResult _backupResult;
        public EObjectSelection _objectSelection;

        // First things first, setup logging 
        private readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #endregion

        #region Constructors

        public caaExport(){}

        public caaExport(IGalaxy Galaxy)
        {
            this.Galaxy = Galaxy;
        }

        #endregion

        #region Properties

        //TODO : Add more sophisticated Get/Set to better sanitize values being set externally

        public IGalaxy Galaxy
        {
            get
            {
                return _galaxy;
            }

            set
            {
                if (value != null)
                {
                    _galaxy = value;
                }
            }
        }

        public bool Connected
        {
            get
            {
                try
                {
                    return (this.Galaxy.CdiVersionString != null);
                }
                catch
                {
                    return false;
                }
            }
        }

        public string GalaxyName
        {
            get
            {
                return _galaxyName;
            }
            set
            {
                _galaxyName  = value;
            }
        }

        public string GRNodeName
        {
            get
            {
                return _grNodeName;
            }
            set
            {
                _grNodeName = value;
            }
        }

        public string Username
        {
            get
            {
                return _username;
            }
            set
            {
                _username = value;
            }
        }

        public string Password
        {
            get
            {
                return _password;
            }
            set
            {
                _password = value;
            }
        }
       
        public string BackupFileName
        {
            get
            {
                return _backupFileName;
            }

            set
            {
                _backupFileName = value;
            }


        }

        public string BackupFolderName
        {
            get
            {
                return _backupFolderName;
            }

            set
            {
                _backupFolderName = value;
            }
        }

        public string DelimitedObjectList
        {
            get
            {
                return _delimitedObjectList;
            }

            set
            {
                _delimitedObjectList = value;
            }
        }

        // TODO: Split backup type into format and selection
        // TODO: Create enumerations for formats and selections
        public string BackupType
        {
            get
            {
                return _backupType;
            }

            set
            {
                _backupType = value;
            }
        }

        public EObjectSelection ObjectSelection
        {
            get
            {
                return _objectSelection;
            }

            set
            {
                _objectSelection = value;
            }
        }


        public EBackupResult  BackupResult
        {
            get
            {
                return _backupResult;
            }

            set
            {
                _backupResult = value;
            }
        }

        public bool IncludeConfigVersion
        {
            get
            {
                return _includeConfigVersion;
            }

            set
            {
                _includeConfigVersion = value;
            }

        }


        //TODO : Use enumerations for FilterType
        public string FilterType
        {
            get
            {
                return _filterType;
            }

            set
            {
                _filterType = value;
            }

        }

        public string Filter
        {
            get
            {
                return _filter;
            }

            set
            {
                _filter = value;
            }
        }

        public DateTime ChangeLogTimestampStartFilter
        {
            get
            {              
                return _changeLogTimestampStartFilter;
            }

            set
            {
                _changeLogTimestampStartFilter = value;
            }
        }

        public string CustomSQLSelection
        {
            get
            {
                return _customSQLSelection;
            }

            set
            {
                _customSQLSelection = value;
            }
        }

        public bool OverwriteFiles
        {
            get
            {
                return _overwriteFiles;
            }

            set
            {
                _overwriteFiles = value;
            }
        }

        public string ObjectListFile
        {
            get
            {
                return _objectListFile;
            }

            set
            {
                _objectListFile = value;
            }
        }

        #endregion

        #region Public Functions

        public int CreateBackup()
        {
            try
            {                 
                this.ExecuteBackup(this.BackupType);
                return 0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Provide a method to set the object selection using a string variable.
        /// </summary>
        /// <param name="ObjectSelection"></param>
        public void SetObjectSelection(String ObjectSelection)
        {
            switch(ObjectSelection.ToLower())
            {
                case "completegalaxy":
                    this.ObjectSelection = EObjectSelection.CompleteGalaxy;
                    break;
              case "allinstances":
                    this.ObjectSelection = EObjectSelection.AllInstances;
                    break;
                case "alltemplates":
                    this.ObjectSelection = EObjectSelection.AllTemplates;
                    break;
                case "filteredobjects":
                    this.ObjectSelection = EObjectSelection.FilteredObjects;
                    break;
                case "objectlist":
                    this.ObjectSelection = EObjectSelection.ObjectList;
                    break;
                default:
                    throw new Exception("Invlid Object Selection "  + ObjectSelection);
            }
        }

        /// <summary>
        /// Provide a method to set the Backup Result as a string
        /// </summary>
        /// <param name="BackupResult"></param>
        public void SetBackupResult(String BackupResult)
        {
            switch(BackupResult.ToLower())
            {
                case "cab":
                    this.BackupResult = EBackupResult.CAB;
                    break;
                case "separateaapkg":
                    this.BackupResult = EBackupResult.SeparateAAPKG;
                    break;
                case "separatecsv":
                    this.BackupResult = EBackupResult.SeparateCSV;
                    break;
                case "singleaapkg":
                    this.BackupResult = EBackupResult.SingleAAPKG;
                    break;
                case "singlecsv":
                    this.BackupResult = EBackupResult.SingleCSV;
                    break;
                default:
                    throw new Exception("Invalid Backup Result " + BackupResult);
            }
        }

        #endregion

        #region Private Functions

        /// <summary>
        /// Get a connection to the Galaxy
        /// </summary>
        /// <returns></returns>
        public bool Connect()
        {
            try
            {
                aaGalaxyHelper agh = new aaGalaxyHelper();
                
                // Attempt a connection
                if(agh.Connect(this.GRNodeName, this.GalaxyName, this.Username, this.Password))
                {
                    // If it is successful then assign the galaxy
                    this.Galaxy = agh.Galaxy;
                }

                // return the result of the connection;
                return this.Connected;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Get a connection to the Galaxy
        /// </summary>
        /// <returns></returns>
        public bool Connect(string GRNodeName, string GalaxyName, string UserName, string Password)
        {
            try
            {
                this.GRNodeName = GRNodeName;
                this.GalaxyName = GalaxyName;
                this.Username = Username;
                this.Password = Password;

                return this.Connect();

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
        private int ExecuteBackup(string BackupType)
        {
            try
            {
                // First check to make sure we have a good conection to a galaxy
                if(!this.Connected)
                {
                    //Try to connect
                    this.Connect();
                }

                // After connection do a basic test to make sure we have a good connection
                try
                {
                    string test = this.Galaxy.CdiVersionString;                   
                }
                catch
                {
                    throw new Exception("Galaxy " + this.GalaxyName + " on " + this.GRNodeName + " not connected.");
                }

                // Instantiate a new object selection class
                cObjectList objectListWorker = new cObjectList(this.Galaxy, this.GRNodeName);
                
                // Create the IGObjects that will be used for backups
                IgObjects objectsToBackup;

                // Set the global filters
                objectListWorker.ChangeLogTimestampStartFilter = this.ChangeLogTimestampStartFilter;
                objectListWorker.CustomSQLSelection = this.CustomSQLSelection;
                
                //Check the ObjectListFile.  If it has been defined then use that as a source of objects
                if (this.ObjectListFile != "")
                {
                    string[] ObjectArray;

                    log.Info("Parsing file " + ObjectListFile + " for the object list");

                    //Test to see if the file exists
                    if (!System.IO.File.Exists(this.ObjectListFile))
                    {
                        throw new Exception(this.ObjectListFile + " does not exist.");
                    }

                    // Read the file into an array.  One line per object
                    ObjectArray = System.IO.File.ReadAllLines(this.ObjectListFile);

                    // If the first line is a CSV then run a split and recreate the array using all the items in the single line
                    if (ObjectArray[0].Contains(','))
                    {
                        ObjectArray = ObjectArray[0].Split(',');
                    }

                    // Now recast back to a single CSV for consumption later.  This will overwrite anything the user has specified on the command line
                    // for an object list
                    this.DelimitedObjectList = string.Join(",", ObjectArray);

                    log.Debug("Parsed the following object list from file " + ObjectListFile + " : " + this.DelimitedObjectList);
                }

                // Check the Object Selection and select the appropriate objects
                switch (this.ObjectSelection)
                {
                    // Nothing to do for this
                    case EObjectSelection.CompleteGalaxy:                            
                        objectsToBackup = null;
                        break;
                            
                    case EObjectSelection.AllInstances:
                        objectsToBackup = objectListWorker.GetAllInstances(true);
                        break;

                    case EObjectSelection.AllTemplates:
                        objectsToBackup = objectListWorker.GetAllTemplates(true);
                        break;

                    case EObjectSelection.ObjectList:
                        objectsToBackup = objectListWorker.GetObjectsFromStringList(this.DelimitedObjectList, true);
                        break;

                    case EObjectSelection.FilteredObjects:
                        objectsToBackup = objectListWorker.GetObjectsFromSingleFilter(this.FilterType, this.Filter, cObjectList.ETemplateOrInstance.Both, true);
                        break;

                    default:
                        // Do Nothing
                        objectsToBackup = null;
                        break;

                }

                // Switch on the object selection again, this time separating out the two major situations which is complete galaxy backup vs item backup
                switch(this.ObjectSelection)
                {
                    // Complete Backups
                    case EObjectSelection.CompleteGalaxy:
                        switch (this.BackupResult)
                        {
                            case EBackupResult.CAB:
                                return BackupCompleteCAB(this.BackupFileName);

                            case EBackupResult.SingleAAPKG:
                                return BackupCompleteAAPKG(this.BackupFileName);
                            
                            case EBackupResult.SingleCSV:
                                return BackupCompleteCSV(this.BackupFileName);

                            default:
                                throw new Exception("Invliad Backup Result " + this.BackupResult.ToString());
                        }

                    // Item based backups
                    case EObjectSelection.AllInstances:
                    case EObjectSelection.AllTemplates:
                    case EObjectSelection.ObjectList:
                    case EObjectSelection.FilteredObjects:
                        switch (this.BackupResult)
                        {
                            case EBackupResult.SingleAAPKG:
                                return BackupToFile(EExportType.exportAsPDF, objectsToBackup, this.BackupFileName);

                            case EBackupResult.SingleCSV:
                                return BackupToFile(EExportType.exportAsCSV, objectsToBackup, this.BackupFileName);

                            case EBackupResult.SeparateAAPKG:
                                return BackupToFolder(EExportType.exportAsPDF, objectsToBackup, this.BackupFolderName);

                            case EBackupResult.SeparateCSV:
                                return BackupToFolder(EExportType.exportAsCSV, objectsToBackup, this.BackupFolderName);

                            default:
                                throw new Exception("Invliad Backup Result " + this.BackupResult.ToString());
                        }

                    default:
                        throw new Exception("Invalid Object Selection " + this.ObjectSelection.ToString());

                }


                //// Determine which type of backup and call the appropriate routine
                //switch (this.BackupType)
                //{
                //    // Call the complete backup routine to generate a CAB
                //    //case "CompleteCAB":
                //    //    return BackupCompleteCAB(this.BackupFileName);

                //    //// Call the complete AAPKG backup routine
                //    //case "CompleteAAPKG":
                //    //    return BackupCompleteAAPKG(this.BackupFileName);

                //    //// Call the complete CSV backup routine
                //    //case "CompleteCSV":
                //    //    return BackupCompleteCSV(this.BackupFileName);



                //    //Export All Templates to an Single AAPKG
                //    case "AllTemplatesAAPKG":
                //        return BackupToFile(EExportType.exportAsPDF, objectListWorker.GetAllTemplates(), this.BackupFileName);

                //    //Export all Instances to Single AAPKG
                //    case "AllInstancesAAPKG":
                //        return BackupToFile(EExportType.exportAsPDF, objectListWorker.GetAllInstances(), this.BackupFileName);

                //    //Export all Instances to Single CSV
                //    case "AllInstancesCSV":
                //        return BackupToFile(EExportType.exportAsCSV, objectListWorker.GetAllInstances(), this.BackupFileName);

                //    //Export All Templates to Separate AAPKG Files
                //    case "AllTemplatesSeparateAAPKG":
                //        return BackupToFolder(EExportType.exportAsPDF, objectListWorker.GetAllTemplates(), this.BackupFolderName);

                //    //Export all Instances to Separate AAPKG
                //    case "AllInstancesSeparateAAPKG":
                //        return BackupToFolder(EExportType.exportAsPDF, objectListWorker.GetAllInstances(), this.BackupFolderName);

                //    //Export all Instances to Separate CSV
                //    case "AllInstancesSeparateCSV":
                //        return BackupToFolder(EExportType.exportAsCSV, objectListWorker.GetAllInstances(), this.BackupFolderName);



                //    //Call the objects AAPKG Backup routine
                //    case "ObjectsSingleAAPKG":
                //        return BackupToFile(EExportType.exportAsPDF, objectListWorker.GetObjectsFromStringList(this.DelimitedObjectList, true), this.BackupFileName);

                //    //Call the objects CSV Backup routine
                //    case "ObjectsSingleCSV":
                //        return BackupToFile(EExportType.exportAsCSV, objectListWorker.GetObjectsFromStringList(this.DelimitedObjectList, true), this.BackupFileName);

                //    //Exporting all Separate objects into AAPKG's
                //    case "ObjectsSeparateAAPKG":
                //        return BackupToFolder(EExportType.exportAsPDF, objectListWorker.GetObjectsFromStringList(this.DelimitedObjectList, true), this.BackupFolderName);

                //    //Exporting all Separate objects into CSV's
                //    case "ObjectsSeparateCSV":
                //        return BackupToFolder(EExportType.exportAsCSV, objectListWorker.GetObjectsFromStringList(this.DelimitedObjectList, true), this.BackupFolderName);


                //    //Export Objects Based on Filter Criteria to single AAPKG
                //    case "FilteredObjectsAAPKG":
                //        return BackupToFile(EExportType.exportAsPDF, objectListWorker.GetObjectsFromSingleFilter(this.FilterType, this.Filter, cObjectList.ETemplateOrInstance.Both, true), this.BackupFileName);

                //    //Export Objects Based on Filter Criteria to single CSV
                //    case "FilteredObjectsCSV":
                //        return BackupToFile(EExportType.exportAsCSV, objectListWorker.GetObjectsFromSingleFilter(this.FilterType, this.Filter, cObjectList.ETemplateOrInstance.Both, true), this.BackupFileName);

                //    //Export Objects Based on Filter Criteria to Separate AAPKG
                //    case "FilteredObjectsSeparateAAPKG":
                //        return BackupToFolder(EExportType.exportAsPDF, objectListWorker.GetObjectsFromSingleFilter(this.FilterType, this.Filter, cObjectList.ETemplateOrInstance.Both, true), this.BackupFolderName);

                //    //Export Objects Based on Filter Criteria to Separate CSV
                //    case "FilteredObjectsSeparateCSV":
                //        return BackupToFolder(EExportType.exportAsCSV, objectListWorker.GetObjectsFromSingleFilter(this.FilterType, this.Filter, cObjectList.ETemplateOrInstance.Both, true), this.BackupFolderName);


                //    default:
                //        throw new Exception("Invalid Backup Type " + this.BackupType);
                //}             
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Perform a complete backup generating a CAB
        /// </summary>
        /// <returns></returns>
        public int BackupCompleteCAB(String BackupFileName)
        {
            int ProcessId;

            try
            {
                // first verify connection
                if (!this.Connected)
                {
                    throw new Exception("Galaxy not connected");
                }

                log.Debug("Checking and correcting filename " + BackupFileName + " to use .CAB");
                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".CAB");

                // Check for file exists.  Bail if the file already exists and we don't want to overwrite
                if (CheckandLogFileExists(BackupFileName)){return 0;}

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
                _galaxy.Backup(ProcessId, BackupFileName, this.GRNodeName, _galaxy.Name);

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
        public int BackupCompleteAAPKG(String BackupFileName)
        {
            //IgObjects workingGObjectList;
            cObjectList ObjectList = new cObjectList(this.Galaxy, this.GRNodeName);

            try
            {
                // first verify connection
                if (!this.Connected)
                {
                    throw new Exception("Galaxy not connected");
                }

                log.Debug("Checking and correcting filename " + BackupFileName + " to use .AAPKG");
                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".AAPKG");

                // Check for file exists.  Bail if the file already exists and we don't want to overwrite
                if (CheckandLogFileExists(BackupFileName)) { return 0; }

                log.Info("Starting AAPKG Backup to " + BackupFileName);

                // Set the filter criteria
                ObjectList.ChangeLogTimestampStartFilter = this.ChangeLogTimestampStartFilter;
                ObjectList.CustomSQLSelection = this.CustomSQLSelection;

                //Perform the export
                ObjectList.GetCompleteObjectList(true).ExportObjects(EExportType.exportAsPDF, BackupFileName);
                
                if (this.Galaxy.CommandResult.Successful != true)
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
        public int BackupCompleteCSV(String BackupFileName)
        {
            //IgObjects workingGObjectList;
            cObjectList ObjectList = new cObjectList(this.Galaxy, this.GRNodeName);

            try
            {

                // first verify connection
                if (!this.Connected)
                {
                    throw new Exception("Galaxy not connected");
                }

                log.Debug("Checking and correcting filename " + BackupFileName + " to use .CSV");

                // Inspect the filename.  Correct the extension if necessary
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, ".CSV");

                // Check for file exists.  Bail if the file already exists and we don't want to overwrite
                if (CheckandLogFileExists(BackupFileName)) { return 0; }

                log.Info("Starting CSV Backup to " + BackupFileName);

                // Set the filter criteria
                ObjectList.ChangeLogTimestampStartFilter = this.ChangeLogTimestampStartFilter;
                ObjectList.CustomSQLSelection = this.CustomSQLSelection;

                ObjectList.GetCompleteObjectList(true).ExportObjects(EExportType.exportAsCSV, BackupFileName);

                if (this.Galaxy.CommandResult.Successful != true)
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
        public int BackupToFile(EExportType ExportType, IgObjects GalaxyObjects, String BackupFileName)
        {
            try
            {
                // first verify connection
                if (!this.Connected)
                {
                    throw new Exception("Galaxy not connected");
                }

                // Make sure we have the right extension on the backup file
                BackupFileName = System.IO.Path.ChangeExtension(BackupFileName, CorrectExtension(ExportType));

                // Check for file exists.  Bail if the file already exists and we don't want to overwrite
                if (CheckandLogFileExists(BackupFileName)) { return 0; }

                log.Info("Backing up to " + BackupFileName);

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
            catch (Exception ex)
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
        public int BackupToFile(EExportType ExportType, IgObject GalaxyObject, String BackupFileName)
        {
            IgObjects GalaxyObjects;
            String[] SingleItemName = new String[1];

            try
            {
                // first verify connection
                if (!this.Connected)
                {
                    throw new Exception("Galaxy not connected");
                }

                SingleItemName[0] = GalaxyObject.Tagname;

                if (SingleItemName[0].Substring(0, 1) == "$")
                {
                    //Template
                    GalaxyObjects = this.Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ref SingleItemName);
                }
                else
                {
                    // Instance
                    GalaxyObjects = this.Galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref SingleItemName);
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
        public int BackupToFolder(EExportType ExportType, IgObjects GalaxyObjects, String BackupFolderName)
        {
            String[] SingleItemName = new String[1];
            IgObjects SingleObject;
            String BackupFileName;
            String Extension;
            String ConfigVersion;

            try
            {
                // first verify connection
                if (!this.Connected)
                {
                    throw new Exception("Galaxy not connected");
                }

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

                //Need to iterate through the the objects
                foreach (IgObject Item in GalaxyObjects)
                {
                    // Populate the single item array with the item's tagname
                    SingleItemName[0] = Item.Tagname;

                    //Default Config Version Text to Blank
                    ConfigVersion = "";

                    //If we require config version in the filename then figure it out then add it
                    if (_includeConfigVersion)
                    {
                        ConfigVersion = "-v" + Item.ConfigVersion.ToString();
                    }

                    // Calculate the appropriate backup filename
                    BackupFileName = BackupFolderName + "\\" + Item.Tagname + ConfigVersion + Extension;

                    // Check for file exists.  Bail if the file already exists and we don't want to overwrite
                    // In this section it is actually a bit redundant but we don't want to perform the Galaxy query if not necessary
                    if (CheckandLogFileExists(BackupFileName)) { continue; }

                    // Figure out if we're dealing with a Template or Instance
                    if (Item.Tagname.Substring(0, 1) == "$")
                    {
                        //Template
                        SingleObject = _galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ref SingleItemName);
                    }
                    else
                    {
                        // Instance
                        SingleObject = _galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref SingleItemName);
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
        private string CorrectExtension(EExportType ExportType)
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

        /// <summary>
        /// Check to see if a file exists.  If it does then log message and return true
        /// </summary>
        /// <param name="Filename"></param>
        /// <returns></returns>
        private bool CheckandLogFileExists(string Filename)
        {
            // Check for file exists.  Bail if the file already exists and we don't want to overwrite
            if (System.IO.File.Exists(Filename) & !this.OverwriteFiles)
            {
                log.Info(Filename + " already exists.  Backup will not be performed");
                return true;
            }
            return false;
        }

        #endregion

        #region Enums

        public enum EObjectSelection
        {
            CompleteGalaxy = 1,
            AllTemplates = 2,
            AllInstances = 3,
            ObjectList = 4,
            FilteredObjects = 5
        }

        public enum EBackupResult
        {
            CAB = 1,
            SingleAAPKG = 2,
            SingleCSV = 3,
            SeparateAAPKG = 4,
            SeparateCSV = 5
        }

        #endregion

    }
}
