using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ArchestrA.GRAccess;
using log4net;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;

namespace aaObjectSelection
{
    
    public class cObjectList
    {
        private IGalaxy _galaxy;
        private string _grNodeName;
        private string _ChangeLogTimestampStartFilter = "";
        private string _CustomSQLSelection = "";
        private SqlConnection _sqlConn = new SqlConnection();

        // First things first, setup logging 
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public cObjectList()
        {
            
        }

        public cObjectList(IGalaxy galaxy)
        {
            _galaxy = galaxy;
        }

        public cObjectList(IGalaxy galaxy, string grNodeName)
        {
            _galaxy = galaxy;
            _grNodeName = grNodeName;
        }

        #region Properties

        public IGalaxy Galaxy
        {
            get
            {
                return _galaxy;
            }

            set
            {
                _galaxy = value;
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

        public string ChangeLogTimestampStartFilter
        {
            get
                {
                    return _ChangeLogTimestampStartFilter;
                }
            set
            {
                _ChangeLogTimestampStartFilter = value;
            }
        }

        public string CustomSQLSelection
        {

            get
            {
                return _CustomSQLSelection;
            }

            set
            {
                _CustomSQLSelection = value;
            }
        }

        #endregion

        #region Core Functions

        /// <summary>
        /// Return an empty set of GObjects so we can simply add to it
        /// </summary>
        /// <returns></returns>
        private IgObjects GetEmptyIgObjects()
        {
            string[] DummyStringRef;
            try
            {
                DummyStringRef = new string[1];
                DummyStringRef[0] = "-";
                return _galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref DummyStringRef);
            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                return null;
            }
        }

        /// <summary>
        /// Get a complete list of objects from galaxy
        /// </summary>
        /// <returns></returns>
        public IgObjects GetCompleteObjectList(bool ApplyFilters = false)
        {
            IgObjects returnGalaxyObjects;

            // Get an empty set of objects to start working with
            returnGalaxyObjects = GetEmptyIgObjects();

            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsTemplate, EConditionType.namedLike, "%"));
            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsInstance, EConditionType.namedLike, "%"));

            if (ApplyFilters)
            {
                this.ApplyCustomFilters(returnGalaxyObjects);
            }

            // Return the complete list
            return returnGalaxyObjects;
        }

        /// <summary>
        /// Get a complete list of all templates from Galaxy
        /// </summary>
        /// <returns></returns>
        public IgObjects GetAllTemplates(bool ApplyFilters = false)
        {
            IgObjects returnGalaxyObjects;

            // Get an empty set of objects to start working with
            returnGalaxyObjects = GetEmptyIgObjects();

            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsTemplate, EConditionType.namedLike, "%"));

            if (ApplyFilters)
            {
                this.ApplyCustomFilters(returnGalaxyObjects);
            }

            // Return the complete list
            return returnGalaxyObjects;
        }

        /// <summary>
        /// Get a complete list of all templates from Galaxy
        /// </summary>
        /// <returns></returns>
        public IgObjects GetAllInstances(bool ApplyFilters = false)
        {
            IgObjects returnGalaxyObjects;

            // Get an empty set of objects to start working with
            returnGalaxyObjects = GetEmptyIgObjects();

            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsInstance, EConditionType.namedLike, "%"));

            if (ApplyFilters)
            {
                this.ApplyCustomFilters(returnGalaxyObjects);
            }

            // Return the complete list
            return returnGalaxyObjects;
        }

        /// <summary>
        /// Get list of objects using a list of strings
        /// </summary>
        /// <param name="ObjectList"></param>
        /// <param name="ApplyFilters"></param>
        /// <returns></returns>
        public IgObjects GetObjectsFromStringList(String ObjectList, bool ApplyFilters = false)
        {
            string[] ObjectArray;
            IgObjects returnGalaxyObjects;

            try
            {
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
                returnGalaxyObjects = _galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ref ObjectArray);

                if (_galaxy.CommandResult.Successful != true)
                {
                    // Failed to retrieve any objects from the query
                    throw new Exception("Error while querying templates by tagname");
                }

                log.Debug("QueryObjectsByName for Instances");
                // Get Instance Objects.  We have to do this in two steps b/c we can't query templates and instances at the same time
                returnGalaxyObjects.AddFromCollection(_galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ref ObjectArray));

                log.Debug("Verify GalaxyObject <> Null, Galaxy Command Success, and Galaxy Object Count > 0");
                if ((returnGalaxyObjects == null) || (_galaxy.CommandResult.Successful != true) || (returnGalaxyObjects.count == 0))
                {
                    // Failed to retrieve any objects from the query
                    throw new Exception("Failed to retrieve objects to export.");
                }

                // Apply the filters if the switch it turned on
                if (ApplyFilters)
                {
                    returnGalaxyObjects = ApplyCustomFilters(returnGalaxyObjects);
                }

                // REturn the Results
                return returnGalaxyObjects;

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Get list of objects using galaxy filters
        /// </summary>
        /// <param name="ObjectList"></param>
        /// <param name="ApplyFilters"></param>
        /// <returns></returns>
        public IgObjects GetObjectsFromSingleFilter(String FilterType, String Filter, ETemplateOrInstance TemplateOrInstance, bool ApplyFilters = false)
        {
            
            IgObjects returnGalaxyObjects;

            try
            {
                // Get an empty set of objects to start working with
                returnGalaxyObjects = GetEmptyIgObjects();

                // Do we need to include templates?
                if ((TemplateOrInstance == ETemplateOrInstance.Template) || (TemplateOrInstance == ETemplateOrInstance.Both))
                {
                    // Now get the template Objects into a GObjects set
                    returnGalaxyObjects.AddFromCollection(_galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ConditionType(FilterType), (object)Filter, EMatch.MatchCondition));
                }

                // Do we need to include instances?
                if ((TemplateOrInstance == ETemplateOrInstance.Instance) || (TemplateOrInstance == ETemplateOrInstance.Both))
                {
                    // Now get the template Objects into a GObjects set
                    returnGalaxyObjects.AddFromCollection(_galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsInstance, ConditionType(FilterType), (object)Filter, EMatch.MatchCondition));
                }

                if ((returnGalaxyObjects == null) || (_galaxy.CommandResult.Successful != true) || returnGalaxyObjects.count == 0)
                {
                    // Failed to retrieve any objects from the query
                    throw new Exception("Failed to retrieve objects to export.");
                }

                // Apply the filters if the switch it turned on
                if (ApplyFilters)
                {
                    returnGalaxyObjects = ApplyCustomFilters(returnGalaxyObjects);
                }

                // REturn the Results
                return returnGalaxyObjects;

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        #endregion

        #region Static Functions

        /// <summary>
        /// Create a delimited string with the object tagnames
        /// </summary>
        /// <param name="GalaxyObjects"></param>
        /// <param name="Delimiter"></param>
        /// <returns></returns>
        public static string AsDelimitedString(IgObjects GalaxyObjects, char Delimiter = ',')
        {
            StringBuilder workingValue = new StringBuilder();
            string returnValue;

            if(GalaxyObjects.count == 0)
            {
                return "";
            }
            
            foreach (IgObject Item in GalaxyObjects)
            {
                // Populate the single item array with the item's tagname
                workingValue.Append(Item.Tagname + Delimiter);
            }

            // Move over to string
            returnValue = workingValue.ToString().TrimEnd(Delimiter);

            // Return final value
            return returnValue;
        }

        #endregion



        #region Utilities

        /// <summary>
        /// Return an EConditionType when given a string that maps to the condition type 
        /// </summary>
        /// <param name="ConditionType"></param>
        /// <returns></returns>
        private EConditionType ConditionType(String ConditionType)
        {
            // Just do a big switch case on all the different kinds, return 
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
        /// Filter the object list based on custom filter information passed
        /// </summary>
        /// <returns></returns>
        private string ApplyCustomFilters(string ObjectList)
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
            if (_ChangeLogTimestampStartFilter != "")
            {
                workingList = GetObjectListForChangeLogAllObjectsAfterTimestampAsCSV(DateTime.Parse(_ChangeLogTimestampStartFilter), ETemplateOrInstance.Both, workingList);
            }

            //Custom SQL
            if (_CustomSQLSelection != "")
            {
                workingList = GetObjectListFromCustomSQL(_CustomSQLSelection, ETemplateOrInstance.Both, workingList);
            }

            //Trim the leading and trailing "
            workingList = workingList.Trim('"');

            returnValue = workingList;

            log.Debug(returnValue);
            return returnValue;
        }

        /// <summary>
        /// Filter the Galaxy object list based on custom filter information passed
        /// </summary>
        /// <returns></returns>
        private IgObjects ApplyCustomFilters(IgObjects GalaxyObjects)
        {
            List<String> galaxyObjectList = new List<String>();
            IgObjects returnGalaxyObjects;
            string workingList = "";
            string[] ObjectArray;
            
            // Get all of the items in the list.  Only good way to do this is 
            foreach (IgObject GObject in GalaxyObjects)
            {
                galaxyObjectList.Add(GObject.Tagname);
            }

            // Set the worklist to the beginning value which is all objects
            workingList = String.Join(",", galaxyObjectList.ToArray());

            // Filter the Original List considering the Timestamp Filter
            if (_ChangeLogTimestampStartFilter != "")
            {
                workingList = GetObjectListForChangeLogAllObjectsAfterTimestampAsCSV(DateTime.Parse(_ChangeLogTimestampStartFilter), ETemplateOrInstance.Both, workingList);
            }

            //Custom SQL
            if (_CustomSQLSelection != "")
            {
                workingList = GetObjectListFromCustomSQL(_CustomSQLSelection, ETemplateOrInstance.Both, workingList);
            }

            //Trim the leading and trailing "
            workingList = workingList.Trim('"');

            // Split the working list in an array of string
            ObjectArray = workingList.Split(',');

            // Get an empty set of objects to start working with
            returnGalaxyObjects = GetEmptyIgObjects();

            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsTemplate, ObjectArray));

            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_galaxy.QueryObjectsByName(EgObjectIsTemplateOrInstance.gObjectIsInstance, ObjectArray));

            // Check for any failures
            if ((returnGalaxyObjects == null) || (_galaxy.CommandResult.Successful != true) || returnGalaxyObjects.count == 0)
            {
                // Failed to retrieve any objects from the query
                throw new Exception("Failed to retrieve objects to export.");
            }

            return returnGalaxyObjects;

        }

        #endregion

        #region SQL Data

        private DataTable GetSQLData(string SQLQuery)
        {
            DataTable returnDataTable;

            try
            {

                // Check the connection
                if (_sqlConn.State != ConnectionState.Open)
                {
                    _sqlConn.ConnectionString = GetSQLConnectionString();
                    _sqlConn.Open();
                }

                if (_sqlConn.State != ConnectionState.Open)
                {
                    throw new Exception("SQL Connection Failed to Open");
                }

                // Setup our command
                SqlCommand sqlCmd = new SqlCommand();

                sqlCmd.Connection = _sqlConn;

                // Set the query text
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.CommandText = SQLQuery;

                // Execute
                returnDataTable = new DataTable("Data");
                returnDataTable.Load(sqlCmd.ExecuteReader());

                return returnDataTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string GetSQLConnectionString()
        {
            // Build the basic connection string
            // For now we force you to use a trusted connection just becuase it is good security hygiene.  Maybe later
            // we will consider letting the caller pass sql creds 
            /// TODO: Let user pass SQL creds to SQL Connection

            return "Server=" + this.GRNodeName + ";Database=" + this.Galaxy.Name +";Trusted_Connection=True;";

        }

        private string GetSQLForChangeLogAllObjectsAfterTimestampAsCSV(DateTime TargetTimestamp, ETemplateOrInstance ItemTypeSelection = ETemplateOrInstance.Both, string ObjectList = "")
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

        private string GetObjectListForChangeLogAllObjectsAfterTimestampAsCSV(DateTime TargetTimestamp, ETemplateOrInstance ItemTypeSelection = ETemplateOrInstance.Both, string ObjectList = "")
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

        private string GetObjectListFromCustomSQL(string SQL, ETemplateOrInstance ItemTypeSelection = ETemplateOrInstance.Both, string ObjectList = "")
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

        #endregion

        #region Enums

        public enum ETemplateOrInstance
        {
            Template = 1,
            Instance = 2,
            Both = 3
        }

        #endregion
    }
}
