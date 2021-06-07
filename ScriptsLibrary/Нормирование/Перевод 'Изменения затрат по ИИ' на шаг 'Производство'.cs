//Перевод 'Изменения затрат по ИИ' на шаг 'Производство'

//На извещении(в конструкторском и технологическом) есть ссылочный атрибут
//«Ссылка на справку об изменении затрат по ИИ БМЗ» (id атрибута 26287).
//Атрибут ссылается на тип объекта «Изменения затрат по ИИ» (id объекта 2796).
//Необходимо:
//В процессе идет ИИ.По этому атрибуту найти этот объект и перевести его
//на шаг производство (id шага 10083). В ИИ может быть и не заполнен этот атрибут.
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using System.Collections.Generic;
using System.Diagnostics;

namespace Scripts
{
    public class Script
    {
        public ICSharpScriptContext ScriptContext { get; set; }

        #region Константы

        private const int costsIIRefAttrID = 26287;//Ссылка на справку об изменении затрат по ИИ БМЗ
        private const int IIType = 1157;
        private const int costsChangesType = 2796;//Изменения затрат по ИИ
        private const int lcStep_prodution = 10083;//создание = 10081, производство = 10083

        #endregion Константы

        public void Execute(IActivity activity)
        {
            if (Debugger.IsAttached)
                Debugger.Break();

            IUserSession UserSession = activity.Session;
            string FinalMessage = string.Empty;

            List<int> IITypes = MetaDataHelper.GetObjectTypeChildrenID(IIType);

            foreach (IAttachment attachment in activity.Attachments)
            {
                if (IITypes.Contains(attachment.ObjectType))
                {
                    //Ссылка на справку об изменении затрат по ИИ БМЗ
                    IDBAttribute costsIIRefAttr = attachment.Object.GetAttributeByID(costsIIRefAttrID);

                    if (costsIIRefAttr == null)
                    {
                        continue;
                    }

                    //Изменение затрат по ИИ
                    IDBObject costsChangingObj = activity.Session.GetObject((long)costsIIRefAttr.Value);

                    if (costsChangingObj.ObjectType == costsChangesType
                        && costsChangingObj.LCStep != lcStep_prodution)
                    {
                        try
                        {
                            //Переводим объект на шаг ЖЦ "Производство"
                            costsChangingObj.LCStep = lcStep_prodution;
                        }
                        catch (NotificationException)
                        {
                            throw;
                        }
                    }
                }
            }
        }
    }
}