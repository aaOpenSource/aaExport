using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ArchestrA.GRAccess;
using Newtonsoft.Json;

namespace aaJSON
{
    public partial class Form1 : Form
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

        #endregion

        public Form1()
        {
            InitializeComponent();
            Setup();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Connect();
            GetObjects();
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

        private static int Connect()
        {
            try
            {

                _GRNodeName = "localhost";
                _GalaxyName = "twtest";
                _Username = "";
                _Password = "";
                
                //log.Debug("Retrieving Galaxies for " + _GRNodeName);
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
                _Galaxy.Login();


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

        private static void GetObjects()
        {
            IgObjects objlist;

            objlist = GetCompleteObjectList();

            string output = JsonConvert.SerializeObject(objlist);

            string objtext;

            int count = objlist.count;
            
            foreach(IgObject obj in objlist)
            {
                objtext = JsonConvert.SerializeObject(obj);
                objtext = obj.Attributes.ToString();
            }

            int x = 1;

        }

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

        private static IgObjects GetCompleteObjectList()
        {
            IgObjects returnGalaxyObjects;

            // Get an empty set of objects to start working with
            returnGalaxyObjects = GetEmptyIgObjects();

            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_Galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsTemplate, EConditionType.namedLike, "%"));
            // Now get the template Objects into a GObjects set
            returnGalaxyObjects.AddFromCollection(_Galaxy.QueryObjects(EgObjectIsTemplateOrInstance.gObjectIsInstance, EConditionType.namedLike, "%"));

            // Return the complete list
            return returnGalaxyObjects;
        }

        private void txtLog_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
