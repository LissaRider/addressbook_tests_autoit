using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace addressbook_tests_autoit
{
    public class GroupHelper : HelperBase
    {
        public static string GROUPWINTITLE = "Group editor";
        public GroupHelper(ApplicationManager manager) : base(manager) { }

        public void Add(GroupData newGroup)
        {
            OpenGroupsDialogue();
            // Нажатие кнопки New
            aux.ControlClick(GROUPWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d53");
            // Ввод значения в поле наименования группы
            aux.Send(newGroup.Name);
            // Эмуляция нажатия кнопки Enter
            aux.Send("{ENTER}");
            CloseGroupsDialogue();
        }

        private void CloseGroupsDialogue()
        {
            // Закрытие окна кнопкой Close
            aux.ControlClick(GROUPWINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d54");
        }

        private void OpenGroupsDialogue()
        {
            // Вызов окна Group editor 
            aux.ControlClick(WINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d512");
            // Ожидание открытия окна
            aux.WinWait(GROUPWINTITLE);
        }

        public List<GroupData> GetGroupList()
        {
            List<GroupData> list = new List<GroupData>();
            // Возвращение пустого списка
            // return new List<GroupData>();
            OpenGroupsDialogue();
            // 4 параметр: команда https://www.autoitscript.com/autoit3/docs/functions/ControlTreeView.htm ,
            // 5 параметр: дополнительный параметр (либо путь, либо порядковый номер),
            // 6 параметр -?
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
    }
}