using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;

namespace Exchange
{
    public class ADUser
    {
        private string result;
        private string resultMessage;

        private string enabledAccount;
        private string exchangeEnabled;
        private string distinguishedName;

        //This is the main email address of the users mailbox
        private string userPrincipalName;
        private string accountPassword;
        private string domain;
        //the ID if the overarching owner of all the users / mailboxes in this account
        private string customerID;
        private string title;
        private string givenName;
        private string surname;
        private string displayName;
        private string description;
        private string manager;
        private string employeeID;
        private string company;
        private string department;
        private string streetAddress;
        private string city;
        private string state;
        private string country;
        private string postcode;
        private string poBox;
        private string office;
        private string officePhone;
        private string homePhone;
        private string mobilePhone;
        private string fax;

        public string getEnabledAccount()
        {
            return enabledAccount;
        }

        public void setEnabledAccount(string enabledAccount)
        {
            this.enabledAccount = enabledAccount;
        }

        public string getExchangeEnabled()
        {
            return exchangeEnabled;
        }

        public void setExchangeEnabled(string exchangeEnabled)
        {
            this.exchangeEnabled = exchangeEnabled;
        }

        public string getDistinguishedName()
        {
            return distinguishedName;
        }

        public void setDistinguishedName(string distinguishedName)
        {
            this.distinguishedName = distinguishedName;
        }

        public string getUserPrincipalName()
        {
            return userPrincipalName;
        }

        public void setUserPrincipalName(string userPrincipalName)
        {
            this.userPrincipalName = userPrincipalName;
        }

        public string getAccountPassword()
        {
            return accountPassword;
        }

        public void setAccountPassword(string accountPassword)
        {
            this.accountPassword = accountPassword;
        }

        public string getDomain()
        {
            return domain;
        }

        public void setDomain(string domain)
        {
            this.domain = domain;
        }

        public string getCustomerID()
        {
            return customerID;
        }

        public void setCustomerID(string customerID)
        {
            this.customerID = customerID;
        }

        public string getTitle()
        {
            return title;
        }

        public void setTitle(string title)
        {
            this.title = title;
        }

        public string getGivenName()
        {
            return givenName;
        }

        public void setGivenName(string givenName)
        {
            this.givenName = givenName;
        }

        public string getSurname()
        {
            return surname;
        }

        public void setSurname(string surname)
        {
            this.surname = surname;
        }

        public string getDisplayName()
        {
            return displayName;
        }

        public void setDisplayName(string displayName)
        {
            this.displayName = displayName;
        }

        public string getDescription()
        {
            return description;
        }

        public void setDescription(string description)
        {
            this.description = description;
        }

        public string getManager()
        {
            return manager;
        }

        public void setManager(string manager)
        {
            this.manager = manager;
        }

        public string getEmployeeID()
        {
            return employeeID;
        }

        public void setEmployeeID(string employeeID)
        {
            this.employeeID = employeeID;
        }

        public string getCompany()
        {
            return company;
        }

        public void setCompany(string company)
        {
            this.company = company;
        }

        public string getDepartment()
        {
            return department;
        }

        public void setDepartment(string department)
        {
            this.department = department;
        }

        public string getStreetAddress()
        {
            return streetAddress;
        }

        public void setStreetAddress(string streetAddress)
        {
            this.streetAddress = streetAddress;
        }

        public string getCity()
        {
            return city;
        }

        public void setCity(string city)
        {
            this.city = city;
        }

        public string getState()
        {
            return state;
        }

        public void setState(string state)
        {
            this.state = state;
        }

        public string getCountry()
        {
            return country;
        }

        public void setCountry(string country)
        {
            this.country = country;
        }

        public string getPostcode()
        {
            return postcode;
        }

        public void setPostcode(string postcode)
        {
            this.postcode = postcode;
        }

        public string getPoBox()
        {
            return poBox;
        }

        public void setPoBox(string poBox)
        {
            this.poBox = poBox;
        }

        public string getOffice()
        {
            return office;
        }

        public void setOffice(string office)
        {
            this.office = office;
        }

        public string getOfficePhone()
        {
            return officePhone;
        }

        public void setOfficePhone(string officePhone)
        {
            this.officePhone = officePhone;
        }

        public string getHomePhone()
        {
            return homePhone;
        }

        public void setHomePhone(string homePhone)
        {
            this.homePhone = homePhone;
        }

        public string getMobilePhone()
        {
            return mobilePhone;
        }

        public void setMobilePhone(string mobilePhone)
        {
            this.mobilePhone = mobilePhone;
        }

        public string getFax()
        {
            return fax;
        }

        public void setFax(string fax)
        {
            this.fax = fax;
        }

        private bool isSuccessful()
        {
            return result != null && result.Equals("Success");
        }

        public string getResult()
        {
            return result;
        }

        public void setResult(string result)
        {
            this.result = result;
        }

        public string getResultMessage()
        {
            return resultMessage;
        }

        public void setResultMessage(string resultMessage)
        {
            this.resultMessage = resultMessage;
        }

        public override string ToString()
        {
            return "ADUser{" +
                    ", result=" + result +
                    ", resultMessage=" + resultMessage +
                    ", enabledAccount=" + enabledAccount +
                    ", exchangeEnabled=" + exchangeEnabled +
                    ", distinguishedName=" + distinguishedName +
                    ", userPrincipalName=" + userPrincipalName +
                    ", accountPassword=" + accountPassword +
                    ", domain=" + domain +
                    ", customerID=" + customerID +
                    ", title=" + title +
                    ", givenName=" + givenName +
                    ", surname=" + surname +
                    ", displayName=" + displayName +
                    ", description=" + description +
                    ", manager=" + manager +
                    ", employeeID=" + employeeID +
                    ", company=" + company +
                    ", department=" + department +
                    ", streetAddress=" + streetAddress +
                    ", city=" + city +
                    ", state=" + state +
                    ", state=" + state +
                    ", country=" + country +
                    ", postcode=" + postcode +
                    ", poBox=" + poBox +
                    ", office=" + office +
                    ", officePhone=" + officePhone +
                    ", homePhone=" + homePhone +
                    ", mobilePhone=" + mobilePhone +
                    ", fax=" + fax +
                    "}\n";
        }
        
        public static ADUser GetAdUser(PSObject adUserObj)
        {
            PSObjectUtils utils = new PSObjectUtils();

            ADUser user = new ADUser();
            user.setResult(utils.GetString(adUserObj, "Result"));
            user.setResultMessage(utils.GetString(adUserObj, "ResultMessage"));
            user.setExchangeEnabled(utils.GetString(adUserObj, "ExchangeEnabled"));
            user.setEnabledAccount(utils.GetString(adUserObj, "EnabledAccount"));
            user.setDistinguishedName(utils.GetString(adUserObj, "DistinguishedName"));
            user.setUserPrincipalName(utils.GetString(adUserObj, "UserPrincipalName"));
            user.setAccountPassword(utils.GetString(adUserObj, "AccountPassword"));
            user.setDomain(utils.GetString(adUserObj, "Domain"));
            user.setCustomerID(utils.GetString(adUserObj, "CustomerID"));
            user.setTitle(utils.GetString(adUserObj, "Title"));
            user.setGivenName(utils.GetString(adUserObj, "GivenName"));
            user.setSurname(utils.GetString(adUserObj, "Surname"));
            user.setDisplayName(utils.GetString(adUserObj, "DisplayName"));
            user.setDescription(utils.GetString(adUserObj, "Description"));
            user.setManager(utils.GetString(adUserObj, "Manager"));
            user.setEmployeeID(utils.GetString(adUserObj, "EmployeeID"));
            user.setCompany(utils.GetString(adUserObj, "Company"));
            user.setDepartment(utils.GetString(adUserObj, "Department"));
            user.setStreetAddress(utils.GetString(adUserObj, "StreetAddress"));
            user.setCity(utils.GetString(adUserObj, "City"));
            user.setState(utils.GetString(adUserObj, "State"));
            user.setCountry(utils.GetString(adUserObj, "Country"));
            user.setPostcode(utils.GetString(adUserObj, "Postcode"));
            user.setPoBox(utils.GetString(adUserObj, "PoBox"));
            user.setOffice(utils.GetString(adUserObj, "Office"));
            user.setOfficePhone(utils.GetString(adUserObj, "OfficePhone"));
            user.setHomePhone(utils.GetString(adUserObj, "HomePhone"));
            user.setMobilePhone(utils.GetString(adUserObj, "MobilePhone"));
            user.setFax(utils.GetString(adUserObj, "Fax"));
            return user;
        }

        public string GetPSParameters()
        {
            return // " -EnabledAccount \"" + enabledAccount + "\"" +
                   // " -ExchangeEnabled \"" + exchangeEnabled + "\"" +
                   // " -DistinguishedName \"" + distinguishedName + "\"" +
                   // " -UserPrincipalName \"" + userPrincipalName + "\"" +
                    " -Identity \"" + userPrincipalName + "\"" +
                    " -AccountPassword \"" + accountPassword + "\"" +
                   // " -Domain \"" + domain + "\"" +
                    " -CustomerID \"" + customerID + "\"" +
                    " -Title \"" + title + "\"" +
                    " -GivenName \"" + givenName + "\"" +
                    " -Surname \"" + surname + "\"" +
                    " -DisplayName \"" + displayName + "\"" +
                    " -Description \"" + description + "\"" +
                   // " -Manager \"" + manager + "\"" +
                   // " -EmployeeID \"" + employeeID + "\"" +
                   // " -Company \"" + company + "\"" +
                    " -Department \"" + department + "\"" +
                    " -StreetAddress \"" + streetAddress + "\"" +
                    " -City \"" + city + "\"" +
                    " -State \"" + state + "\"" +
                    " -Country \"" + country + "\"" +
                   // " -Postcode \"" + postcode + "\"" +
                    " -PoBox \"" + poBox + "\"" +
                   // " -Office \"" + office + "\"" +
                    " -OfficePhone \"" + officePhone + "\"" +
                   // " -HomePhone \"" + homePhone + "\"" +
                    " -MobilePhone \"" + mobilePhone + "\"" +
                   // " -Fax \"" + fax + "\"";
                   "";
        }
    }
}
