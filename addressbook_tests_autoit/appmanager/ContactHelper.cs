using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace addressbook_tests_autoit
{
    public class ContactHelper : HelperBase
    {
        public static string CONTACTEDITORWINTITLE = "Contact Editor";
        public static string QUESTIONWINTITLE = "Question";

        public ContactHelper(ApplicationManager manager) : base(manager)
        {
        }

        public ContactHelper Create(ContactData contact)
        {
            InitContactCreation();
            FillContactForm(contact);
            SubmitContactCreationAndCloseDialogue();
            return this;
        }

        public ContactHelper Remove(int index)
        {
            SelectContact(index);
            InitContactRemoval();
            SubmitContactRemoval();
            return this;
        }

        // Contact creation methods
        public void InitContactCreation()
        {
            // Function 'ControlClick' https://www.autoitscript.com/autoit3/docs/functions/ControlClick.htm
            // ControlClick("title", "text", controlID)
            // Sends a mouse click command to a given control.
            aux.ControlClick(WINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d58");
            // Function 'WinWait' https://www.autoitscript.com/autoit3/docs/functions/WinWait.htm
            // WinWait("title", "text", "timeout(seconds)")
            // Pauses execution of the script until the requested window exists.
            aux.WinWait(CONTACTEDITORWINTITLE);
        }

        public void FillContactForm(ContactData contact)
        {
            // Function 'ControlFocus' https://www.autoitscript.com/autoit3/docs/functions/ControlFocus.htm
            // ControlFocus("title", "text", controlID)
            // Sets input focus to a given control on a window.
            aux.ControlFocus(CONTACTEDITORWINTITLE, "", "WindowsForms10.EDIT.app.0.2c908d516");
            // Function 'Send' https://www.autoitscript.com/autoit3/docs/functions/Send.htm
            // Send("keys")
            // Sends simulated keystrokes to the active window.
            aux.Send(contact.Firstname);
            aux.ControlFocus(CONTACTEDITORWINTITLE, "", "WindowsForms10.EDIT.app.0.2c908d513");
            aux.Send(contact.Lastname);
        }

        public void SubmitContactCreationAndCloseDialogue()
        {
            aux.ControlClick(CONTACTEDITORWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d58");
        }

        // Common methods for contacts
        private void SelectContact(int index)
        {
            // Function 'ControlListView' https://www.autoitscript.com/autoit3/docs/functions/ControlListView.htm
            // ControlListView("title", "text", controlID, "command 'Select'", "", "")
            // Selects one or more items.
            aux.ControlListView(WINTITLE, "", "WindowsForms10.Window.8.app.0.2c908d510",
                "Select", index.ToString(), "");
        }

        // Contact removal methods
        private void InitContactRemoval()
        {
            aux.ControlClick(WINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d59");
            aux.WinWait(QUESTIONWINTITLE);
        }

        private void SubmitContactRemoval()
        {
            aux.ControlClick(QUESTIONWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d52");
        }

        // Verification
        public List<ContactData> GetContactsList()
        {
            List<ContactData> list = new List<ContactData>();
            int count = GetContactCount();

            for (int i = 0; i < count; i++)
            {
                // Function 'ControlListView' https://www.autoitscript.com/autoit3/docs/functions/ControlListView.htm
                // ControlListView ( "title", "text", controlID, "command 'GetText'", Item, SubItem)
                // Returns the text of a given item/subitem.
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

        public int GetContactCount()
        {
            // Function 'ControlListView' https://www.autoitscript.com/autoit3/docs/functions/ControlListView.htm
            // ControlListView("title", "text", controlID, "command 'GetItemCount'", "", "")
            // Returns the number of list items.
            return int.Parse(aux.ControlListView(
                WINTITLE, "", "WindowsForms10.Window.8.app.0.2c908d510",
                "GetItemCount", "", ""));
        }
    }
}