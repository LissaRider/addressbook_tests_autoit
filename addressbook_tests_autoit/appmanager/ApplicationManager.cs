using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoItX3Lib;

namespace addressbook_tests_autoit
{   
    public class ApplicationManager
    {
        public static string WINTITLE = "Free Address Book";
        private AutoItX3 aux;
        private GroupHelper groupHelper;

        // Запуск приложения с помощью команды Run()
        public ApplicationManager()
        {
            aux = new AutoItX3();
            aux.Run(@"C:\Program Files (x86)\GAS Softwares\Free Address Book\AddressBook.exe", "", aux.SW_SHOW);
            aux.WinWait(WINTITLE);
            aux.WinActivate(WINTITLE);
            aux.WinWaitActive(WINTITLE);

            groupHelper = new GroupHelper(this);
        }

        // Остановка приложения с помощью команды ControlClick()
        public void Stop()
        {
            // 1 параметр: Название окна, в котором находится кнопка,
            // 2 параметр: уточняющий параметр - текст кнопки (необязательный),
            // 3 параметр: идентификатор кнопки - локатор.
            aux.ControlClick(WINTITLE, "", "WindowsForms10.BUTTON.app.0.2c908d510");
        }

        public AutoItX3 Aux
        {
            get
            {
                return aux;
            }
        }

        public GroupHelper Groups
        {
            get
            {
                return groupHelper;
            }
        }
    }
}
