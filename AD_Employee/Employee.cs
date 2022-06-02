using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.Text.RegularExpressions;
using System.DirectoryServices.AccountManagement;
namespace AD_Employee
{
    public class Employee
    {

        public bool changeTitle(string employeeIDstring, string newTitle)
        {
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);
            if (!isValidPropertyData(newTitle)) throw new System.ArgumentException("invalid title", newTitle);

            if (newTitle == lookupOnePropertyValue(employeeIDstring, "title")) return false; // no change necessary

            if (changeOneProperty(employeeIDstring, "title", newTitle) &&
                changeOneProperty(employeeIDstring, "description", newTitle)) return true;

            throw new System.InvalidOperationException("failed: changeTitle(" + employeeIDstring + "," + newTitle + ")");
        }

        public bool changeCompany(string employeeIDstring, string company)
        {
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);
            if (!isValidPropertyData(company)) throw new System.ArgumentException("invalid company", company);

            if (company == lookupOnePropertyValue(employeeIDstring, "company")) return false; // no change necessary

            if (changeOneProperty(employeeIDstring, "company", company)) return true;

            throw new System.InvalidOperationException("failed: changeCompany(" + employeeIDstring + "," + company + ")");
        }

        public bool changeOffice(string employeeIDstring, string office)
        {
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);
            if (!isValidPropertyData(office)) throw new System.ArgumentException("invalid office", office);

            if (office == lookupOnePropertyValue(employeeIDstring, "physicaldeliveryofficename")) return false; // no change necessary

            if (changeOneProperty(employeeIDstring, "physicaldeliveryofficename", office)) return true;

            throw new System.InvalidOperationException("failed: changeOffice(" + employeeIDstring + "," + office + ")");
        }

        public bool changeLocation(string employeeIDstring, string division, string street, string city, string state, string zipcode)
        {
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);
            if (!isValidPropertyData(division)) throw new System.ArgumentException("invalid division", division);
            //if (!isValidPropertyData(street)) throw new System.ArgumentException("invalid street", street);
            //if (!isValidPropertyData(city)) throw new System.ArgumentException("invalid city", city);
            //if (!isValidPropertyData(state)) throw new System.ArgumentException("invalid state", state);
            //if (!isValidPropertyData(zipcode)) throw new System.ArgumentException("invalid zipcode", zipcode);

            if ((division == lookupOnePropertyValue(employeeIDstring, "division")))
                //&& (street == lookupOnePropertyValue(employeeIDstring, "streetaddress"))
                //&& (city == lookupOnePropertyValue(employeeIDstring, "l"))
                //&& (state == lookupOnePropertyValue(employeeIDstring, "st""))
                //&& (zipcode == lookupOnePropertyValue(employeeIDstring, postalcode")))
                return false; // no change necessary

            if (changeOneProperty(employeeIDstring, "division", division))
                //&& changeOneProperty(employeeIDstring, "streetaddress", street) &&
                //changeOneProperty(employeeIDstring, "l", city) &&
                //changeOneProperty(employeeIDstring, "st", state) &&
                //changeOneProperty(employeeIDstring, "postalcode", zipcode))
                return true;

            throw new System.InvalidOperationException
                ("failed: changeLocation(" + employeeIDstring + "," + division + "," + street + "," + city + "," + state + "," + zipcode + ")");
        }

        public bool changeOU(string employeeIDstring, string OUstring)
        {
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);
                     

            if (changeOneProperty(employeeIDstring, "OU", OUstring)) return true;



            throw new System.InvalidOperationException("failed: changeManager(" + employeeIDstring + ")");
        }

        public bool changeManager(string employeeIDstring, string managerIDstring)
        {
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);

            string managerDistinguishedName = lookupOnePropertyValue(managerIDstring, "distinguishedname");
            if (managerDistinguishedName == lookupOnePropertyValue(employeeIDstring, "manager")) return false; // no change necessary

            if (changeOneProperty(employeeIDstring, "manager", managerDistinguishedName)) return true;

            throw new System.InvalidOperationException("failed: changeManager(" + employeeIDstring + "," + managerIDstring + ")");
        }

        public bool nameChangeNecessary(string employeeIDstring, string newName)
        {
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);
            if (!isValidPropertyData(newName)) throw new System.ArgumentException("invalid new name", newName);

            if (newName == lookupOnePropertyValue(employeeIDstring, "cn")) return false; // no change necessary

            return true;
        }

        public bool isValidEmployeeID(string employeeIDstring)
        {
            // regular expression match string for employeeID:
            // "JVF" or "J8S" followed by exactly six digits
            string validEmployeeID = @"^(HGR)[0-9]{6}$";
            string validMallardEmployeeID = @"^(Q13)[0-9]{6}$";
            string validTorrentEmployeeID = @"^(TES)[0-9]{6}$";

            if (Regex.IsMatch(employeeIDstring, validEmployeeID)) return true;
            if (Regex.IsMatch(employeeIDstring, validMallardEmployeeID)) return true;
            if (Regex.IsMatch(employeeIDstring, validTorrentEmployeeID)) return true;
            else return false;
        }

        public bool isValidPropertyData(string employeePropertyData)
        {
            // regular expression match string for properties/attribute data:
            // letters, commas, apostrophes, hyphens, parentheses and whitespace
            // length of 1 to 64 characters
            string toMatchRegEx = @"^[A-Za-z0-9,\.\/'\-\(\)\s]{1,64}$";

            if (Regex.IsMatch(employeePropertyData, toMatchRegEx)) return true;
            else return false;
        }

        public bool isValidPropertyDataLong(string employeePropertyData)
        {
            // regular expression match string for properties/attribute data:
            // letters, commas, apostrophes, hyphens, parentheses and whitespace
            // length of 1 to 64 characters
            string toMatchRegEx = @"^[A-Za-z0-9,=\.'\-\(\)\s]{1,256}$";

            if (Regex.IsMatch(employeePropertyData, toMatchRegEx)) return true;
            else return false;
        }


        public string lookupOnePropertyValue(string employeeIDstring, string propertyName)
        {
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);

            try
            {
                DirectoryEntry myLdapConnection = createDirectoryEntry();
                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search.Filter = "(employeeID=" + employeeIDstring + ")";
                SearchResult result = search.FindOne();
                if (result == null) return null; // throw new System.InvalidOperationException("employeeID not found: " + employeeIDstring);

                // employee exists, loop through LDAP fields
                ResultPropertyCollection fields = result.Properties;
                foreach (String ldapField in fields.PropertyNames)
                {
                    if (ldapField == propertyName)
                    {
                        string returnString = "";
                        foreach (Object myCollection in fields[ldapField])
                            returnString += myCollection.ToString();
                        return returnString;
                    }
                }
                return null; // throw new System.InvalidOperationException("propertyName " + propertyName + " not found for employee " + employeeIDstring);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        // private functions

        private bool changeOneProperty(string employeeIDstring, string propertyToChange, string newValue)
        {
            if (!isValidEmployeeID(employeeIDstring)) throw new System.ArgumentException("invalid employeeID", employeeIDstring);
            if (propertyToChange == "manager")
            {
                if (!isValidPropertyDataLong(newValue)) throw new System.ArgumentException("invalid attribute value", newValue);
            }
            else
            {
                if (!isValidPropertyData(newValue)) throw new System.ArgumentException("invalid attribute value", newValue);
            }

            try
            {
                DirectoryEntry myLdapConnection = createDirectoryEntry();
                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search.Filter = "(employeeID=" + employeeIDstring + ")";
                search.PropertiesToLoad.Add(propertyToChange);

                SearchResult result = search.FindOne();
                if (result == null) throw new System.InvalidOperationException("employee not found: " + employeeIDstring);

                DirectoryEntry entryToUpdate = result.GetDirectoryEntry();
                entryToUpdate.Properties[propertyToChange].Value = newValue;
                entryToUpdate.CommitChanges();
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private DirectoryEntry createDirectoryEntry()
        {
            // create and return new LDAP connection
            DirectoryEntry ldapConnection = new DirectoryEntry("bayouwellservices.com");
            ldapConnection.Path = "LDAP://DC=bayouwellservices,DC=com";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            return ldapConnection;


        }

        private DirectoryEntry createDirectoryEntry(string strGroup,string strLOB)
        {
            // create and return new LDAP connection

            string strState = string.Empty;


            if (strLOB == "Wireline Production")
            {
                strState = "Patriot Well";
            }

            else if (strGroup == "Bowie" || strGroup == "Wharton" || strGroup == "Houston" || strGroup == "Pleasanton" || strGroup == "Midland" || strGroup == "Odessa" || strGroup == "Big Spring" || strGroup == "Kermit" || strGroup == "Kilgore" || strGroup == "Snyder" || strGroup == "Andrews" || strGroup == "Denver City")
            {
                strState = "Texas";
            }
             else if (strGroup == "New Town" ||  strGroup == "Dickinson"  || strGroup == "Belfield" )
            {
                strState = "North Dakota";
            }
            else if (strGroup == "Milliken" || strGroup == "Brighton" || strGroup == "Hobbs" || strGroup == "Fort Morgan")
            {
                strState = "Colorado";
            }
            else if (strGroup == "Hobbs" || strGroup == "Artesia")
            {
                strState = "New Mexico";
            }
            else if (strGroup == "Casper" || strGroup == "Gillette")
            {
                strState = "Wyoming";
            }




            DirectoryEntry ldapConnection = new DirectoryEntry("bayouwellservices.com");
            ldapConnection.Path = "LDAP://OU=Users,OU="+ strGroup +",OU="+strState+",DC=bayouwellservices,DC=com";
        //    ldapConnection.Path = "LDAP://OU=Users,OU=Midland,OU=Patriot Well,DC=bayouwellservices,DC=com";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            return ldapConnection;


        }
        public static void SetProperty(DirectoryEntry de, string PropertyName, string PropertyValue)
        {
            if (PropertyValue != null)
            {
                if (de.Properties.Contains(PropertyName))
                {
                    de.Properties[PropertyName][0] = PropertyValue;
                }
                else
                {
                    de.Properties[PropertyName].Add(PropertyValue);
                }
            }
        }


        // create new user from workflow
        public string TermUser(string employeeID)
        {
            try
            {
                DataSet1TableAdapters.ADPEmployeesTableAdapter EmpTa = new DataSet1TableAdapters.ADPEmployeesTableAdapter();
                DataSet1TableAdapters.EmployeesTableAdapter TicketEmpTa = new DataSet1TableAdapters.EmployeesTableAdapter();
                //set status to term ?????
               // int i = TicketEmpTa.UpdateNoAllowance("Active", title, location, fullName, Convert.ToDecimal(rate), type, Convert.ToDecimal(rate2), profitCenter, locationCode, LOB, "", employeeID);
               // i = EmpTa.UpdateNoAllowance(location, profitCenter, fullName, title, "", "Active", supervisor, email, locationCode, LOBCode, deptCode, dept, LOB, employeeID);
               
                          
                DisableandMoveAccount(employeeID);              
                return "ok";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }



        // create new user from workflow
        public string CreateNewUser(string employeeID, string fullName,  string email, string location,string LOB,string profitCenter,string title,double rate,string type,double rate2,DateTime hireDate,string supervisor,string dept)
        {
            try
            {

              

                Boolean bAllowance = false;
                decimal dAllownace = 0;
                decimal dRate3 = 0;
                //decimal dRate = 0;
                //decimal dRate2 = 0;
                string deptCode = string.Empty;
                string LOBCode = string.Empty;
                string LOBName = string.Empty;
                string locationCode = string.Empty;
                string locationName = string.Empty;
                string login = string.Empty;
                string lastName = string.Empty;
                string strRate = string.Empty;
                string ADName = string.Empty;

                List<string> names = fullName.Split(' ').ToList();
                string firstName = names.First();
                names.RemoveAt(0);
                if (names.Count > 1)
                {
                    lastName = string.Join(" ", names);
                }
                else
                {
                    lastName = names[0].ToString();
                }

                fullName = lastName + "," + firstName;
                ADName = firstName + ' ' + lastName;
                strRate = rate.ToString();
                login = email.Substring(0,email.LastIndexOf('@'));
                supervisor = supervisor.Substring(supervisor.LastIndexOf('-') + 1);

                //if (rate != "")
                //{
                //   dRate = Convert.ToDecimal(rate);

                //}
                //else
                //{
                //    dRate = 0;
                //}

                //if (rate2 != "")
                //{
                //    dRate2 = Convert.ToDecimal(rate2);

                //}
                //else
                //{
                //    dRate2 = 0;
                //}


                 deptCode = LOB.Substring(0, 2);
                 LOBCode = LOB.Substring(0, 2);
                 LOBName = LOB.Substring(3);      
               
                locationCode = location.Substring(0,3);
                locationName = location.Substring(4);

                DataSet1TableAdapters.ADPEmployeesTableAdapter EmpTa = new DataSet1TableAdapters.ADPEmployeesTableAdapter();
       
                DataSet1TableAdapters.EmployeesTableAdapter TicketEmpTa = new DataSet1TableAdapters.EmployeesTableAdapter();
            




                //save employee to database for reporting
                int i = TicketEmpTa.UpdateNoAllowance("Active", title, location, fullName, Convert.ToDecimal(rate), type, Convert.ToDecimal(rate2), profitCenter, locationCode, LOB, "", employeeID);

                if (i < 1)
                {
                    TicketEmpTa.InsertEmployee(fullName, employeeID, "Active", title, locationCode, Convert.ToDecimal(rate), type, Convert.ToDecimal(rate2), profitCenter, location, LOB, dRate3, dAllownace, "");
                }


                i = EmpTa.UpdateNoAllowance(location, profitCenter, fullName, title, "", "Active", supervisor, email, locationCode, LOBCode, deptCode, dept, LOB, employeeID);

                if (i < 1)
                {
                    EmpTa.Insert(employeeID, location, profitCenter, fullName, title, hireDate.ToShortDateString(), "", "Active", supervisor, email, locationCode, LOBCode, deptCode, dept, LOB, bAllowance);
                }


                DirectoryEntry de = createDirectoryEntry(locationName, LOBName);
            
                de.Username = "jpadmin";
                de.Password = "Pwnzor93";
                /// 1. Create user account
                DirectoryEntries users = de.Children;
                DirectoryEntry newuser = users.Add("CN=" + ADName, "user");
                /// 2. Set properties
                SetProperty(newuser, "employeeID", employeeID);
                SetProperty(newuser, "givenname", firstName);
                SetProperty(newuser, "sn", lastName);
                SetProperty(newuser, "SAMAccountName", login);
                SetProperty(newuser, "userPrincipalName", email);
                SetProperty(newuser, "mail", email);
                SetProperty(newuser, "title", title);
                newuser.CommitChanges();
                //SetPW(login, "R@ng3r12");
                SetPassword(newuser.Path);
                newuser.CommitChanges();
                /// 4. Enable account
               EnableAccount(newuser);
              

                newuser.Close();
                de.Close();
                return "ok";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }



        public  void SetPW(string userName, string Password)
        {
            PrincipalContext context = new PrincipalContext(ContextType.Domain);
            
           context.ValidateCredentials("jpadmin", "Pwnzor93");
            UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, userName);
         
            //Enable Account if it is disabled
            user.Enabled = true;
            //Reset User Password
            user.SetPassword(Password);
            //Force user to change password at next logon
            user.ExpirePasswordNow();
            user.Save();
        }

        public void SetPassword(string path)
        {
            try
            {
                DirectoryEntry usr = new DirectoryEntry();
                usr.Username = "jpadmin";
                usr.Password = "Pwnzor93";
                usr.Path = path;
                usr.AuthenticationType = AuthenticationTypes.Secure;
                object[] password = new object[] { "R@ng3r12" };
                object ret = usr.Invoke("SetPassword", password);
                usr.CommitChanges();
                usr.Close();
            }
            catch (Exception e)
            {
               throw e;
            }
        }

        //public void AddToGroup(string userDn, string groupDn)
        //{
        //    try
        //    {

        //        DirectoryEntry dirEntry = createDirectoryEntry(groupDn);

        //        dirEntry.Properties["user"].Add(userDn);
        //        dirEntry.Invoke("Add", new object[] {  userDn });


        //        dirEntry.CommitChanges();
        //        dirEntry.Close();
        //    }
        //    catch (System.DirectoryServices.DirectoryServicesCOMException E)
        //    {
        //        doSomething with E.Message.ToString();

        //    }
        //}

        private static void EnableAccount(DirectoryEntry de)
        {
            try
            {
                //UF_DONT_EXPIRE_PASSWD 0x10000
                de.Username = "jpadmin";
                de.Password = "Pwnzor93";
                int exp = (int)de.Properties["userAccountControl"].Value;
                de.Properties["userAccountControl"].Value = exp | 0x0001;
                de.CommitChanges();
                //UF_ACCOUNTEnable  ~0x0002
                int val = (int)de.Properties["userAccountControl"].Value;
                de.Properties["userAccountControl"].Value = val & ~0x0002;
                de.CommitChanges();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private  void DisableandMoveAccount(string employeeIDstring)
        {
            try
            {
                DirectoryEntry myLdapConnection = createDirectoryEntry();
                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search.Filter = "(employeeID=" + employeeIDstring + ")";
             

                SearchResult result = search.FindOne();
                if (result == null) throw new System.InvalidOperationException("employee not found: " + employeeIDstring);

                DirectoryEntry de = result.GetDirectoryEntry();
              
            
                


                DirectoryEntry de2 = new DirectoryEntry("LDAP://OU=Disabled Users - Waiting for Deletion,DC=bayouwellservices,DC=com");
              
                
                de.Username = "jpadmin";
                de.Password = "Pwnzor93";
                //UF_ACCOUNTDISABLE 0x0002
                int val = (int)de.Properties["userAccountControl"].Value;
                de.Properties["userAccountControl"].Value = val | 0x0002;
                de.CommitChanges();

                de.MoveTo(de2);
                de2.CommitChanges();
                de2.Close();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

   



    }
}
