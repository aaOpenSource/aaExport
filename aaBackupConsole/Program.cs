using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Threading;
using System.Diagnostics;

using ArchestrA.GRAccess;
using log4net;

using Classes.ObjectList;
using Classes.Encryption;
using Classes.Backup;


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
                //Connect();

                // Instantiate the Object and set the parameters
                var BackupObj = new aaBackup()
                {
                    Galaxy = _Galaxy,
                    GRNodeName = _GRNodeName,
                    GalaxyName = _GalaxyName,
                    Username = _Username,
                    Password = _Password,
                    BackupFileName = _BackupFileName,
                    BackupFolderName = _BackupFolderName,
                    DelimitedObjectList = _ObjectList,
                    BackupType = _BackupType,
                    IncludeConfigVersion = (_IncludeConfigVersion == "true") || (_IncludeConfigVersion == "1"),
                    FilterType = _FilterType,
                    Filter = _Filter,
                    ChangeLogTimestampStartFilter = DateTime.Parse(_ChangeLogTimestampStartFilter),
                    CustomSQLSelection = _CustomSQLSelection                    
                };

                // Execute the backup
                BackupObj.CreateBackup();

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
                //_GRAccess = new GRAccessApp();

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
                CheckAndSetParameters(ref _ChangeLogTimestampStartFilter, "ChangeLogTimestampStartFilter", CommandLine, true, "1/1/1970");
                CheckAndSetParameters(ref _CustomSQLSelection, "CustomSQLSelection", CommandLine, true,"");

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

        #endregion

    }
}
