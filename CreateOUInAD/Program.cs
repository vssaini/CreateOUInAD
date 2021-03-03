using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CreateOUInAD
{
    class Program
    {
        static void Main()
        {
            // Credentials to active directory
            const string domain = "domain.com";
            const string user = "Administrator";
            const string password = "Pass99";

            // Possible distinguished names provided by end user 
            //const string dn = "CN=Test User,OU=NewOU,DC=Domain,DC=Com"; // dn of user
            const string dn = "CN=Test User,OU=Men,OU=Staff,OU=NewOU,DC=Domain,DC=Com"; // destinationOU

            // Get domain path as per DC
            var domainPath = GetDomainPath(domain);

            // Retrieve all OUs from path
            var orgUnits = GetAllOUFromPath(dn);

            // Get container
            Console.WriteLine("The destination for users will be {0}", GetDestinationOUPath(domainPath, orgUnits));
            Console.WriteLine();

            orgUnits.Reverse(); 
            for (var i = 0; i < orgUnits.Count; i++)
            {
                var orgUnit = orgUnits[i];
                var ouPath = GetOUPath(domainPath, null, orgUnits, i);

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Creating OU '{0}' in container '{1}' ...", orgUnit, ouPath);
                Console.WriteLine("------------------------------------------------");
                CreateOrganizationalUnit(domain, user, password, orgUnit, ouPath);
            }

            // Hold console so that we can read message
            Console.ReadKey();
        }

        /// <summary>
        /// Get domain path splitted as per DC format.
        /// </summary>
        /// <param name="domain">The domain name</param>
        public static string GetDomainPath(string domain)
        {
            //var sbPath = new StringBuilder();

            //// Split domain name
            //var dCs = domain.Split('.');

            //// Add DCs
            //foreach (var part in dCs)
            //    sbPath.AppendFormat("DC={0},", part);

            //// Remove last "," character and return path
            ////return sbPath.ToString().TrimEnd(',');

            var domainLdap = string.Empty;

            try
            {
                DirectoryContext context = new DirectoryContext(DirectoryContextType.Domain, domain);
                Domain objDomain = Domain.GetDomain(context);
                DirectoryEntry de = objDomain.GetDirectoryEntry();
                domainLdap = Convert.ToString(de.Properties["DistinguishedName"].Value);
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {

            }
            return domainLdap;
        }

        /// <summary>
        /// Get all OU values in string.
        /// </summary>
        /// <param name="path">The path to search for ou.</param>
        /// <returns>Return list of OUs.</returns>
        public static List<string> GetAllOUFromPath(string path)
        {
            const string pattern = @"OU=(.*?),";
            var regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            var matches = regex.Matches(path);

            return (from Match match in matches
                    select match.Groups[1].Value).ToList();
        }

        /// <summary>
        /// Check if object exists in AD.
        /// </summary>
        /// <param name="directoryEntry">The directory entry to check for path</param>
        /// <returns>Return true if exists else false.</returns>
        public static bool IsOUExist(DirectoryEntry directoryEntry)
        {
            bool exist;

            // Validate with Guid
            try
            {
                // ReSharper disable once UnusedVariable
                //var tmp = directoryEntry.Guid; // Had to use this if access to AD needs credentials
                exist = DirectoryEntry.Exists(directoryEntry.Path);
            }
            catch (Exception)
            {
                exist = false;
            }

            return exist;
        }

        /// <summary>
        /// Create organizational unit in active directory.
        /// </summary>
        /// <param name="domain">The doamin name</param>
        /// <param name="user">The user having permission to active directory</param>
        /// <param name="password">The password for connecting to active directory.</param>
        /// <param name="orgUnit">The name of organizational unit</param>
        /// <param name="container">The container under which to create OU.</param>
        public static void CreateOrganizationalUnit(string domain, string user, string password, string orgUnit, string container)
        {
            // Set OU name
            var ouName = string.Format("OU={0}", orgUnit);

            try
            {
                //Create LDAP access to domain
                using (var parentEntry = new DirectoryEntry(domain, user, password))
                {
                    parentEntry.Path = container;

                    // Create OU
                    using (var newOU = parentEntry.Children.Add(ouName, "OrganizationalUnit"))
                    {
                        if (!IsOUExist(newOU))
                        {
                            newOU.CommitChanges();

                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine("OU '{0}' created successfully!", orgUnit);
                            Console.WriteLine();
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Error - OU '{0}' already exists!", orgUnit);
                            Console.WriteLine();
                        }

                    }
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error :- " + e);
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Create LDAP path  to root container under which to create OU.
        /// </summary>
        /// <param name="domainPath">The domain name</param>
        /// <param name="domainController">The domain controller name</param>
        /// <param name="orgUnits">The list of OUs</param>
        /// <param name="ouCounter">The counter to which need to run loop</param>
        /// <returns>Return full ldap path.</returns>
        public static string GetOUPath(string domainPath, string domainController, List<string> orgUnits, int ouCounter)
        {
            // 1st - LDAP://DC1.domain.com/DC=domain,DC=com;
            // 2nd - LDAP://DC1.domain.com/OU=NewOU,DC=domain,DC=com
            // 3rd - LDAP://DC1.domain.com/OU=Staff,OU=NewOU,DC=domain,DC=com;

            // Create string builder and initialize with LDAP Provider
            var sbPath = new StringBuilder("LDAP://");

            // Add domain controller if available
            if (!string.IsNullOrEmpty(domainController))
            {
                sbPath.Append(domainController);
                sbPath.Append("/");
            }

            // Append OUs if available as per condition
            if (ouCounter > 0)
            {
                // Get only that much OU that are required for container
                var selectedOu = new List<string>();
                for (var i = 0; i < ouCounter; i++)
                {
                    selectedOu.Add(orgUnits[i]);
                }

                // Now reverse it to create path in correct LDAP way
                selectedOu.Reverse();
                foreach (var ou in selectedOu)
                {
                    sbPath.AppendFormat("OU={0},", ou);
                }
            }

            // Append domain path
            sbPath.Append(domainPath);

            // Return path
            return sbPath.ToString();

        }

        /// <summary>
        /// Get destination OU path.
        /// </summary>
        /// <param name="domainPath">The LDAP domain path</param>
        /// <param name="orgUnits">The list of organizational units.</param>
        /// <returns>Return container or destination OU LDAP path.</returns>
        public static string GetDestinationOUPath(string domainPath, List<string> orgUnits)
        {
            var sbPath = new StringBuilder();
            foreach (var ou in orgUnits)
            {
                sbPath.AppendFormat("OU={0},", ou);
            }

            // Append domain path
            sbPath.Append(domainPath);

            // Return path
            return sbPath.ToString();
        }
    }
}
