using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ArchestrA.GRAccess;
using log4net;

namespace Classes.GalaxyHelper
{
    class aaGalaxyHelper
    {

        // First things first, setup logging 
       private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
       
        private IGalaxy _galaxy;

        #region properties

        public IGalaxy Galaxy
        {
            get
            {
                return this._galaxy;
            }

            set
            {
                this._galaxy = value;
            }


        }

        public bool Connected
        {
            get
            {
                if (this._galaxy == null)
                {
                    return false;
                }

                return (this._galaxy.CdiVersionString != null);
            }

        }

        #endregion

        /// <summary>
        /// Attempt to make a connection to the Galaxy
        /// </summary>
        /// <returns></returns>
        public bool Connect(string GRNodeName, string GalaxyName, string UserName, string Password)
        {

            GRAccessApp grAccess;
            IGalaxy galaxy;
            IGalaxies galaxies;
    
            try
            {
                // Instantiate the GR Access App
                grAccess = new GRAccessApp();

                log.Debug("Retrieving Galaxies for " + GRNodeName);

                // Get a list of the available galaxies
                galaxies = grAccess.QueryGalaxies(GRNodeName);

                log.Debug("Getting Galaxy Reference for " + GalaxyName);

                //Get a reference to the Galaxy
                galaxy = galaxies[GalaxyName];

                log.Debug("Checking for Success");

                // Check to make sure we have a good reference to the Galaxy
                if (galaxy == null || !grAccess.CommandResult.Successful)
                {
                    throw new Exception(GalaxyName + " is not a legal Galaxy on " + GRNodeName);
                }

                log.Debug("Logging in with Username " + UserName);

                // Attempt to Login
                galaxy.LoginEx(UserName, Password,false);

                log.Debug("Checking for login success");

                //Check for success
                if (!galaxy.CommandResult.Successful)
                {
                    throw new Exception("Galaxy Login Failed");                    
                }

                // If we made it to hear we're good!
                log.Info("Login Succeeded");

                // Set the local reference
                this.Galaxy = galaxy;

                // return the connected status
                return this.Connected;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
