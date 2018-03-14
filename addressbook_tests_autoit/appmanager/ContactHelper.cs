using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace addressbook_tests_autoit
{
    public class ContactHelper : HelperBase
    {
        public static string CONTACTEDITORWINTITLE = "Contact Editor";
        public static string QUESTIONWINTITLE = "Question";

        public ContactHelper(ApplicationManager manager) : base(manager) { }

        public void Create(ContactData contact)
        {
            InitContactCreation();
            FillContactForm(contact);
            SaveAndClose();
        }

        public void Remove(int index)
        {
            SelectContact(index);
            InitContactRemoval();
            SubmitContactRemove();
        }           

        // Contact creation methods
        public void InitContactCreation()
        {
            aux.ControlClick(WINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d58");
            aux.WinWait(CONTACTEDITORWINTITLE);
        }

        public void FillContactForm(ContactData contact)
        {
            aux.ControlFocus(CONTACTEDITORWINTITLE, "", "WindowsForms10.EDIT.app.0.2c908d516");
            aux.Send(contact.Firstname);
            aux.ControlFocus(CONTACTEDITORWINTITLE, "", "WindowsForms10.EDIT.app.0.2c908d513");
            aux.Send(contact.Lastname);
        }

        public void SaveAndClose()
        {
            aux.ControlClick(CONTACTEDITORWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d58");
        }

        // Common methods https://www.autoitscript.com/autoit3/docs/functions/ControlListView.htm
        public void SelectContact(int index)
        {
            aux.ControlListView(WINTITLE, "", "WindowsForms10.Window.8.app.0.2c908d510", "Select", index.ToString(), "");
        }

        // Contact removal methods
        public void InitContactRemoval()
        {
            aux.ControlClick(WINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d59");
            aux.WinWait(QUESTIONWINTITLE);
        }

        public void SubmitContactRemove()
        {
            aux.ControlClick(QUESTIONWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d52");
        }

        public List<ContactData> GetContactsList()
        {
            List<ContactData> list = new List<ContactData>();

            int count = GetContactCount();

            for (int i = 0; i < count; i++)
            {
                string firstName = aux.ControlListView(
                    WINTITLE, "", "WindowsForms10.Window.8.app.0.2c908d510",
                    "GetText", i.ToString(), "0");
                string lastName = aux.ControlListView(
                    WINTITLE, "", "WindowsForms10.Window.8.app.0.2c908d510",
                    "GetText", i.ToString(), "1");
                list.Add(new ContactData()
                {
                    Firstname = firstName,
                    Lastname = lastName
                });
            }
            return list;
        }

        public int GetContactCount()
        {
            return int.Parse(aux.ControlListView(
                WINTITLE, "", "WindowsForms10.Window.8.app.0.2c908d510",
                "GetItemCount", "", ""));
        }

        // Verification
        public void VerifyContactPresence()
        {
            if (GetContactCount() < 1)
            {
                ContactData newContact = new ContactData()
                {
                    Firstname = "FirstName",
                    Lastname = "LastName"
                };

                Create(newContact);
            }
        }
    }
}