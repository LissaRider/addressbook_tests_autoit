using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace addressbook_tests_autoit
{
    public class GroupHelper : HelperBase
    {
        public static string GROUPWINTITLE = "Group editor";
        public static string DELETEGROUPWINTITLE = "Delete group";

        public GroupHelper(ApplicationManager manager) : base(manager)
        {
        }

        public GroupHelper Create(GroupData group)
        {
            OpenGroupsDialogue();

            InitGroupCreation();
            // Ввод значения в поле наименования группы
            aux.Send(group.Name);
            // Эмуляция нажатия кнопки 'Enter'
            aux.Send("{ENTER}");

            CloseGroupsDialogue();
            return this;
        }


        public GroupHelper Remove(int index)
        {
            OpenGroupsDialogue();

            SelectGroup(index);
            InitGroupRemoval();
            SubmitGroupRemoval();

            CloseGroupsDialogue();
            return this;
        }

        // Common methods for groups   
        public void OpenGroupsDialogue()
        {
            // Вызов окна 'Group editor' 
            aux.ControlClick(WINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d512");
            // Ожидание открытия окна 'Group editor'
            aux.WinWait(GROUPWINTITLE);
        }

        public void CloseGroupsDialogue()
        {
            // Закрытие окна 'Group editor' кнопкой 'Close'
            aux.ControlClick(GROUPWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d54");
        }

        public void SelectGroup(int index)
        {
            // Выбор группы из списка
            aux.ControlTreeView(GROUPWINTITLE, "", "WindowsForms10.SysTreeView32.app.0.2c908d51", "Select", "#0|#" + index, "");
        }

        // Group creation methods
        public void InitGroupCreation()
        {
            // Нажатие кнопки 'New'
            aux.ControlClick(GROUPWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d53");
        }

        // Group removal methods
        public void InitGroupRemoval()
        {
            aux.ControlClick(GROUPWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d51");
            aux.WinWait(DELETEGROUPWINTITLE);
        }

        public void SubmitGroupRemoval()
        {
            aux.ControlClick(DELETEGROUPWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d53");
        }

        public List<GroupData> GetGroupsList()
        {
            List<GroupData> list = new List<GroupData>();

            OpenGroupsDialogue();

            // 4 параметр: Команда https://www.autoitscript.com/autoit3/docs/functions/ControlTreeView.htm ,
            // 5 параметр: Дополнительный параметр (либо путь, либо порядковый номер),
            // 6 параметр: Options (необязательный).
            string count = aux.ControlTreeView(
                GROUPWINTITLE, "", "WindowsForms10.SysTreeView32.app.0.2c908d51",
                "GetItemCount", "#0", "");
            for (int i = 0; i < int.Parse(count); i++)
            {
                string item = aux.ControlTreeView(
                    GROUPWINTITLE, "", "WindowsForms10.SysTreeView32.app.0.2c908d51",
                    "GetText", "#0|#" + i, "");
                list.Add(new GroupData()
                {
                    Name = item
                });
            }
            CloseGroupsDialogue();
            return list;
        }

        // Verification
        public void VerifyGroupPresence()
        {
            OpenGroupsDialogue();

            if (GetGroupCount() <= 1)
            {
                GroupData newGroup = new GroupData()
                {
                    Name = "GroupName"
                };
                Create(newGroup);
            }
        }
        public int GetGroupCount()
        {
            OpenGroupsDialogue();

            int count = int.Parse(aux.ControlTreeView(
                GROUPWINTITLE, "", "WindowsForms10.SysTreeView32.app.0.2c908d51",
                "GetItemCount", "#0", ""));

            CloseGroupsDialogue();
            return count;
        }
    }
}