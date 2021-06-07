//Проверка заявок на техподдержку
using Intermech;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    public class TechSupportErrors
    {
        //Необходимо при запуске процесса проверить наличие вложения у данного процесса.
        //Если нет вложения выдать сообщение о том, что необходимо приложить заявку к процессу.
        //Так же ограничить количество вложений, по одному процессу может быть запущен только 1 объект.
        //Если вложений больше одного выдать сообщение о том, что по процессу не разрешено запускать более одной заявки.
        //Для отработки, объект «
        //на техподдержку».

        public ICSharpScriptContext ScriptContext { get; private set; }

        public void Execute(IActivity activity)
        {
            if (activity.Attachments.Count == 1 && activity.Attachments[0].ObjectType != 2481/*Заявка на техподдержку*/)
                throw new NotificationException(String.Format("Во вложении должен быть объект «Заявка на техподдержку»."));

            if (activity.Attachments.Count == 0)
                throw new NotificationException(String.Format("Необходимо добавить заявку во вложения процесса."));

            if (activity.Attachments.Count > 1)
                throw new NotificationException(String.Format("Во вложении должена быть только одна заявка."));

            int techSupType = activity.Session.GetObjectType("Заявка на техподдержку").ObjectType;
        }
    }

    //Необходима проверка наличия и заполнения атрибутов у объекта.
    //Объект тот же, «Заявка на техподдержку». Необходимо 2 скрипта, 
    //т.к.в двух разных местах проверить надо.
    
    //В первом скрипте необходимые
    //атрибуты для проверки: «Вид заявки на техподдержку», «Приоритет заявки на техподдержку»,
    //«Вид ПО (для заявки на техподдержку)», «Характеристика заявки на техподдержку»,
    //«Тема заявки на техподдержку», «Описание заявки на техподдержку».

    //Во втором скрипте: «Причина обращения(заявка на техподдержку)»,
    //«Описание выполнения заявки на техподдержку», «Дата выполнения заявки на техподдержку».

    //Если атрибуты не заполнены или нет, выдавать сообщение о том, что необходимо заполнить такой-то атрибут.

    public class TechSupportAttrErrors1
    {

        public ICSharpScriptContext ScriptContext { get; private set; }

        public void Execute(IActivity activity)
        {
            string error = string.Empty;

            int techSupType = activity.Session.GetObjectType("Заявка на тех. поддержку").ObjectType;
            foreach (IAttachment item in activity.Attachments)
            {
                IDBObject techSup = item.Object;

                IDBAttribute atr_type = techSup.GetAttributeByGuid(new Guid("cbb5d225-d85f-4152-94ce-6c388164db12" /*Вид заявки на техподдержку*/));
                IDBAttribute atr_type2 = techSup.GetAttributeByGuid(new Guid("4d859d67-a11e-4275-9f10-cffa74848762" /*Вид задачи БМЗ*/));

                IDBAttribute atr_prior = techSup.GetAttributeByGuid(new Guid("5f0005e7-7ece-4324-aad5-a5675965c699" /*Приоритет заявки на техподдержку*/));

                IDBAttribute atr_softType = techSup.GetAttributeByGuid(new Guid("ccc33697-55c6-4446-8a68-647f6544f1e8" /*Вид ПО (для заявки на техподдержку)*/));

                IDBAttribute atr_desc = techSup.GetAttributeByGuid(new Guid("aabff3e2-b063-4c7c-8475-ad4246df062d" /*Характеристика заявки на техподдержку*/));
                IDBAttribute atr_desc2 = techSup.GetAttributeByGuid(new Guid("41352d21-870a-498f-818b-9286d1f4f34f" /*Характеристика задачи БМЗ*/));

                IDBAttribute atr_theme = techSup.GetAttributeByGuid(new Guid("6b0c9adf-96c2-48e5-8fbd-1f693f02a17f" /*Тема заявки на техподдержку*/));
                IDBAttribute atr_theme2 = techSup.GetAttributeByGuid(new Guid("3a9dfa69-da91-4df7-a558-e9f797c08563" /*Обозначение+Тема заявки на техподдержку*/));

                IDBAttribute atr_description = techSup.GetAttributeByGuid(new Guid("688e4a74-d301-483d-b180-59c511bfe837" /*Описание заявки на техподдержку*/));
                IDBAttribute atr_description2 = techSup.GetAttributeByGuid(new Guid("caea0f23-8775-4197-9873-c3cfa59fb59e" /*Описание задачи БМЗ*/));

                if (item.ObjectType == techSupType)
                {

                    if (!((atr_type != null) || (atr_type2 != null)))
                        error += string.Format("У объекта {0} не найден атрибут *Вид заявки на техподдержку*\n\r", techSup.NameInMessages);

                    if (atr_prior == null)
                        error += string.Format("У объекта {0} не найден атрибут *Приоритет заявки на техподдержку*\n\r", techSup.NameInMessages);

                    if (atr_softType == null)
                        error += string.Format("У объекта {0} не найден атрибут *Вид ПО (для заявки на техподдержку)*\n\r", techSup.NameInMessages);

                    if (!((atr_desc != null) || (atr_desc2 != null)))
                        error += string.Format("У объекта {0} не найден атрибут *Характеристика заявки на техподдержку*\n\r", techSup.NameInMessages);

                    if (!((atr_theme != null) || (atr_theme2 != null)))
                        error += string.Format("У объекта {0} не найден атрибут *Тема заявки на техподдержку*\n\r", techSup.NameInMessages);

                    if (!((atr_description != null) || (atr_description2 != null)))
                        error += string.Format("У объекта {0} не найден атрибут *Описание заявки на техподдержку*\n\r", techSup.NameInMessages);

                }

            }

            if (error.Length > 0)
                throw new ClientException(error);
        }
    }

    public class TechSupportAttrErrors2
    {
        public ICSharpScriptContext ScriptContext { get; private set; }

        public void Execute(IActivity activity)
        {
            string error = string.Empty;

            int techSupType = activity.Session.GetObjectType("Заявка на тех. поддержку").ObjectType;
            foreach (IAttachment item in activity.Attachments)
            {
                IDBObject techSup = item.Object;

                IDBAttribute atr_reason = techSup.GetAttributeByGuid(new Guid("a4485b08-8e6e-4f98-9449-53b57e96902e" /*Причина обращения (заявка на техподдержку)*/));
                IDBAttribute atr_descript = techSup.GetAttributeByGuid(new Guid("c3794f20-7a0a-46e1-803e-34886d289fe0" /*Описание выполнения заявки на техподдержку*/));
                IDBAttribute atr_date = techSup.GetAttributeByGuid(new Guid("f82e6632-ff62-46ae-b21b-f615f88fb4c6" /*Дата выполнения заявки на техподдержку*/));

                if (item.ObjectType == techSupType)
                {
                    if (atr_reason == null)
                        error += string.Format("У объекта {0} не найден атрибут *Причина обращения (заявка на техподдержку)*\n\r", techSup.NameInMessages);

                    if (atr_descript == null)
                        error += string.Format("У объекта {0} не найден атрибут *Описание выполнения заявки на техподдержку*\n\r", techSup.NameInMessages);

                    if (atr_date == null)
                        error += string.Format("У объекта {0} не найден атрибут *Дата выполнения заявки на техподдержку*\n\r", techSup.NameInMessages);
                }
            }

            if (error.Length > 0)
                throw new NotificationException(error);
        }
    }
}
